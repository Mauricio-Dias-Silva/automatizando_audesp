import os
import tempfile
import shutil
import zipfile
import re
from django.shortcuts import render
from django.http import HttpResponse, FileResponse, HttpResponseBadRequest
from django.views.decorators.http import require_http_methods
from automatizando_core import criar_csv_rascunho, processar_final

# optional Cloud Vision integration
try:
    from automatizando_gcs import upload_file_to_bucket, async_ocr_pdf_to_local
except Exception:
    upload_file_to_bucket = None
    async_ocr_pdf_to_local = None

# optional PDF parsing
try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None


def index(request):
    return render(request, 'main/index.html')


@require_http_methods(["POST"])
def create_csv_view(request):
    files = request.FILES.getlist('xml_files')
    if not files:
        return HttpResponseBadRequest('Nenhum XML enviado')

    temp_dir = tempfile.mkdtemp(prefix='audesp_xml_')
    try:
        # save uploaded files, and extract XMLs from PDFs when provided
        for f in files:
            filename = f.name
            dest = os.path.join(temp_dir, filename)
            with open(dest, 'wb') as out:
                for chunk in f.chunks():
                    out.write(chunk)

            # if PDF, prefer Cloud Vision OCR (if enabled), otherwise fallback
            if filename.lower().endswith('.pdf'):
                used_cloud = False
                try:
                    bucket = os.environ.get('GCS_BUCKET')
                    use_cv = os.environ.get('USE_CLOUD_VISION', '0')
                    if use_cv == '1' and upload_file_to_bucket and async_ocr_pdf_to_local and bucket:
                        # upload and async OCR, results saved to temp_dir
                        dest_blob = f'uploads/{int(time.time())}_{filename}'
                        upload_file_to_bucket(bucket, dest, dest_blob)
                        async_ocr_pdf_to_local(bucket, dest_blob, temp_dir)
                        used_cloud = True
                except Exception:
                    used_cloud = False

                if not used_cloud:
                    try:
                        extract_xmls_from_pdf(dest, temp_dir)
                    except Exception:
                        pass

        csv_path, count = criar_csv_rascunho(temp_dir)
        with open(csv_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type='text/csv; charset=utf-8')
            response['Content-Disposition'] = f'attachment; filename="{os.path.basename(csv_path)}"'
            return response
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


@require_http_methods(["POST"])
def process_view(request):
    csv_file = request.FILES.get('csv_file')
    xml_files = request.FILES.getlist('xml_files')

    if not csv_file or not xml_files:
        return HttpResponseBadRequest('É necessário enviar o CSV preenchido e os XMLs originais')

    temp_dir = tempfile.mkdtemp(prefix='audesp_process_')
    xml_dir = os.path.join(temp_dir, 'xmls')
    os.makedirs(xml_dir, exist_ok=True)

    try:
        csv_path = os.path.join(temp_dir, csv_file.name)
        with open(csv_path, 'wb') as out:
            for chunk in csv_file.chunks():
                out.write(chunk)

        # save uploaded xml/pdf files and extract xmls from pdfs
        for f in xml_files:
            dest = os.path.join(xml_dir, f.name)
            with open(dest, 'wb') as out:
                for chunk in f.chunks():
                    out.write(chunk)

            if f.name.lower().endswith('.pdf'):
                used_cloud = False
                try:
                    bucket = os.environ.get('GCS_BUCKET')
                    use_cv = os.environ.get('USE_CLOUD_VISION', '0')
                    if use_cv == '1' and upload_file_to_bucket and async_ocr_pdf_to_local and bucket:
                        dest_blob = f'uploads/{int(time.time())}_{f.name}'
                        upload_file_to_bucket(bucket, dest, dest_blob)
                        async_ocr_pdf_to_local(bucket, dest_blob, xml_dir)
                        used_cloud = True
                except Exception:
                    used_cloud = False

                if not used_cloud:
                    try:
                        extract_xmls_from_pdf(dest, xml_dir)
                    except Exception:
                        pass

        # processar_final irá criar pasta SAIDA_AUDESP_PRONTOS dentro de xml_dir
        sucesso = processar_final(csv_path, xml_dir)
        saida_dir = os.path.join(xml_dir, 'SAIDA_AUDESP_PRONTOS')
        if not os.path.isdir(saida_dir):
            return HttpResponse('Processado, mas não foi gerada a pasta de saída.', status=500)

        # zipar a saída
        zip_path = os.path.join(temp_dir, 'saida_audesp.zip')
        shutil.make_archive(zip_path.replace('.zip', ''), 'zip', saida_dir)

        # If configured, upload output zip (or directory) to GCS and return link
        upload_output = os.environ.get('UPLOAD_OUTPUT_TO_GCS', '0')
        bucket = os.environ.get('GCS_BUCKET')
        if upload_output == '1' and bucket:
            try:
                # upload whole directory under outputs/<timestamp>/
                out_prefix = os.environ.get('OUTPUT_PREFIX', f'outputs/{int(time.time())}')
                uploaded = []
                if 'upload_directory_to_bucket' in globals():
                    # prefer helper if available
                    from automatizando_gcs import upload_directory_to_bucket, generate_signed_url
                    uploaded = upload_directory_to_bucket(bucket, saida_dir, out_prefix.rstrip('/') + '/SAIDA_AUDESP_PRONTOS')
                    # generate signed url for zip
                    dest_blob = f"{out_prefix.rstrip('/')}/SAIDA_AUDESP_PRONTOS/saida_audesp.zip"
                    signed = generate_signed_url(bucket, dest_blob)
                    html = f'<p>Processado. <a href="{signed}">Baixar saída (zip)</a></p>'
                    return HttpResponse(html)
            except Exception:
                pass

        with open(zip_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type='application/zip')
            response['Content-Disposition'] = 'attachment; filename="SAIDA_AUDESP_PRONTOS.zip"'
            return response

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def extract_xmls_from_pdf(pdf_path, dest_dir):
    """Try extracting embedded XML files from a PDF.
    1) If the PDF has attachments, extract those ending with .xml
    2) Otherwise, extract text and search for an XML snippet starting with '<?xml'
    Writes any found XML files into `dest_dir`.
    """
    if PdfReader is None:
        return

    try:
        reader = PdfReader(pdf_path)
    except Exception:
        return

    # 1) attachments (if supported by pypdf)
    try:
        attachments = getattr(reader, 'attachments', None)
        if attachments:
            for name, data in attachments.items():
                if name.lower().endswith('.xml'):
                    out_path = os.path.join(dest_dir, name)
                    with open(out_path, 'wb') as fh:
                        fh.write(data)
            return
    except Exception:
        pass

    # 2) fallback: extract text and try to find XML snippet
    try:
        text = []
        for p in reader.pages:
            t = p.extract_text()
            if t:
                text.append(t)
        text = '\n'.join(text)
        if '<?xml' in text:
            start = text.find('<?xml')
            # try to find closing tag - naive: last occurrence of '</'
            end = text.rfind('>')
            xml_snippet = text[start:end+1]
            # sanitize filename
            base = os.path.basename(pdf_path)
            name = os.path.splitext(base)[0] + '.xml'
            out_path = os.path.join(dest_dir, name)
            with open(out_path, 'w', encoding='utf-8') as fh:
                fh.write(xml_snippet)
    except Exception:
        return
