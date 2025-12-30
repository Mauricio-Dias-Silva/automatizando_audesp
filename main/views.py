import os
import tempfile
import shutil
import zipfile
from django.shortcuts import render
from django.http import HttpResponse, FileResponse, HttpResponseBadRequest
from django.views.decorators.http import require_http_methods
from automatizando_core import criar_csv_rascunho, processar_final


def index(request):
    return render(request, 'main/index.html')


@require_http_methods(["POST"])
def create_csv_view(request):
    files = request.FILES.getlist('xml_files')
    if not files:
        return HttpResponseBadRequest('Nenhum XML enviado')

    temp_dir = tempfile.mkdtemp(prefix='audesp_xml_')
    try:
        for f in files:
            dest = os.path.join(temp_dir, f.name)
            with open(dest, 'wb') as out:
                for chunk in f.chunks():
                    out.write(chunk)

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

        for f in xml_files:
            dest = os.path.join(xml_dir, f.name)
            with open(dest, 'wb') as out:
                for chunk in f.chunks():
                    out.write(chunk)

        # processar_final irá criar pasta SAIDA_AUDESP_PRONTOS dentro de xml_dir
        sucesso = processar_final(csv_path, xml_dir)
        saida_dir = os.path.join(xml_dir, 'SAIDA_AUDESP_PRONTOS')
        if not os.path.isdir(saida_dir):
            return HttpResponse('Processado, mas não foi gerada a pasta de saída.', status=500)

        # zipar a saída
        zip_path = os.path.join(temp_dir, 'saida_audesp.zip')
        shutil.make_archive(zip_path.replace('.zip', ''), 'zip', saida_dir)

        with open(zip_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type='application/zip')
            response['Content-Disposition'] = 'attachment; filename="SAIDA_AUDESP_PRONTOS.zip"'
            return response

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
