import os
import time
import json
import tempfile
from google.cloud import storage
from google.cloud import vision_v1 as vision


def upload_file_to_bucket(bucket_name, source_file_path, destination_blob_name):
    client = storage.Client()
    bucket = client.bucket(bucket_name)
    blob = bucket.blob(destination_blob_name)
    blob.upload_from_filename(source_file_path)
    return f'gs://{bucket_name}/{destination_blob_name}'


def download_blob_to_file(bucket_name, blob_name, dest_path):
    client = storage.Client()
    bucket = client.bucket(bucket_name)
    blob = bucket.blob(blob_name)
    blob.download_to_filename(dest_path)


def download_prefix_to_dir(bucket_name, prefix, dest_dir):
    """Download all blobs under prefix into dest_dir preserving filenames."""
    client = storage.Client()
    blobs = client.list_blobs(bucket_name, prefix=prefix)
    paths = []
    for blob in blobs:
        if blob.name.endswith('/'):
            continue
        filename = os.path.basename(blob.name)
        out_path = os.path.join(dest_dir, filename)
        blob.download_to_filename(out_path)
        paths.append(out_path)
    return paths


def upload_directory_to_bucket(bucket_name, local_dir, dest_prefix):
    """Upload all files from local_dir to bucket under dest_prefix."""
    client = storage.Client()
    bucket = client.bucket(bucket_name)
    uploaded = []
    for root, _, files in os.walk(local_dir):
        for fname in files:
            local_path = os.path.join(root, fname)
            rel = os.path.relpath(local_path, local_dir)
            blob_name = f"{dest_prefix.rstrip('/')}/{rel}"
            blob = bucket.blob(blob_name)
            blob.upload_from_filename(local_path)
            uploaded.append(f'gs://{bucket_name}/{blob_name}')
    return uploaded


def generate_signed_url(bucket_name, blob_name, expiration_seconds=3600):
    """Generate a signed URL for a blob valid for expiration_seconds."""
    client = storage.Client()
    bucket = client.bucket(bucket_name)
    blob = bucket.blob(blob_name)
    url = blob.generate_signed_url(expiration=expiration_seconds)
    return url


def async_ocr_pdf_to_local(bucket_name, gcs_source_blob_name, local_dest_dir, output_prefix='vision_output/'):
    """Uploads already assumed; gcs_source_blob_name is path inside bucket (no gs://)
    Calls Vision asyncBatchAnnotateFiles on the PDF and downloads resulting JSONs, extracts text and writes any found XML snippets to local_dest_dir.
    Returns list of local xml file paths.
    """
    storage_client = storage.Client()
    vision_client = vision.ImageAnnotatorClient()

    gcs_source_uri = f'gs://{bucket_name}/{gcs_source_blob_name}'
    gcs_output_uri = f'gs://{bucket_name}/{output_prefix}'

    # Configure request
    from google.cloud.vision_v1 import AsyncAnnotateFileRequest, InputConfig, Feature, OutputConfig, GcsSource, GcsDestination

    gcs_source = GcsSource(uri=gcs_source_uri)
    input_config = InputConfig(gcs_source=gcs_source, mime_type='application/pdf')

    gcs_destination = GcsDestination(uri=gcs_output_uri)
    output_config = OutputConfig(gcs_destination=gcs_destination)

    feature = Feature(type_=Feature.Type.DOCUMENT_TEXT_DETECTION)
    request = AsyncAnnotateFileRequest(features=[feature], input_config=input_config, output_config=output_config)

    operation = vision_client.async_batch_annotate_files(requests=[request])
    op = operation
    op.result(timeout=600)

    # list output files in bucket prefix
    bucket = storage_client.bucket(bucket_name)
    blobs = list(storage_client.list_blobs(bucket_name, prefix=output_prefix))

    xml_paths = []
    for blob in blobs:
        if not blob.name.endswith('.json'):
            continue
        # download JSON to temp
        tmpf = tempfile.NamedTemporaryFile(delete=False, suffix='.json')
        tmpf.close()
        blob.download_to_filename(tmpf.name)
        try:
            with open(tmpf.name, 'r', encoding='utf-8') as fh:
                data = json.load(fh)
            # responses -> fullTextAnnotation -> text
            for resp in data.get('responses', []):
                fta = resp.get('fullTextAnnotation')
                if not fta:
                    continue
                text = fta.get('text', '')
                if '<?xml' in text:
                    start = text.find('<?xml')
                    xml_text = text[start:]
                    # write to file
                    base = os.path.basename(gcs_source_blob_name)
                    name = os.path.splitext(base)[0] + '_' + os.path.basename(blob.name).replace('.json', '.xml')
                    out_path = os.path.join(local_dest_dir, name)
                    with open(out_path, 'w', encoding='utf-8') as out:
                        out.write(xml_text)
                    xml_paths.append(out_path)
        except Exception:
            pass
        finally:
            try:
                os.remove(tmpf.name)
            except Exception:
                pass

    return xml_paths
