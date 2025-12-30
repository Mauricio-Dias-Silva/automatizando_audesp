from django.core.management.base import BaseCommand, CommandError
import tempfile
import os
import shutil
from automatizando_gcs import download_blob_to_file, download_prefix_to_dir, upload_directory_to_bucket
from automatizando_core import processar_final

class Command(BaseCommand):
    help = 'Run processing job reading CSV and XMLs from GCS and uploading results back to GCS'

    def add_arguments(self, parser):
        parser.add_argument('--bucket', required=True, help='GCS bucket name')
        parser.add_argument('--csv-blob', required=True, help='GCS blob path for the CSV (e.g. path/PLANILHA.csv)')
        parser.add_argument('--xmls-prefix', required=True, help='GCS prefix where XMLs are located (e.g. uploads/xmls/)')
        parser.add_argument('--output-prefix', required=True, help='GCS prefix to upload outputs into (e.g. outputs/2025-12-30/)')

    def handle(self, *args, **options):
        bucket = options['bucket']
        csv_blob = options['csv_blob']
        xmls_prefix = options['xmls_prefix']
        output_prefix = options['output_prefix']

        tmp = tempfile.mkdtemp(prefix='gcs_job_')
        try:
            # download CSV
            csv_local = os.path.join(tmp, os.path.basename(csv_blob))
            download_blob_to_file(bucket, csv_blob, csv_local)

            # download xmls under prefix
            xmls_dir = os.path.join(tmp, 'xmls')
            os.makedirs(xmls_dir, exist_ok=True)
            download_prefix_to_dir(bucket, xmls_prefix, xmls_dir)

            # process
            count = processar_final(csv_local, xmls_dir)
            self.stdout.write(self.style.SUCCESS(f'Processed {count} items'))

            # upload outputs directory
            saida_dir = os.path.join(xmls_dir, 'SAIDA_AUDESP_PRONTOS')
            if os.path.isdir(saida_dir):
                uploaded = upload_directory_to_bucket(bucket, saida_dir, output_prefix.rstrip('/') + '/SAIDA_AUDESP_PRONTOS')
                for u in uploaded:
                    self.stdout.write(u)
            else:
                self.stdout.write('No output dir created')

        except Exception as e:
            raise CommandError(str(e))
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
