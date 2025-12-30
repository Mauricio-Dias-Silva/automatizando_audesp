from django.core.management.base import BaseCommand, CommandError
import argparse
import logging
from automatizando_core import criar_csv_rascunho, processar_final

logger = logging.getLogger('run_script')
logger.setLevel(logging.INFO)

class Command(BaseCommand):
    help = 'Run automatizando_audesp tasks (create CSV or process CSV and generate XMLs)'

    def add_arguments(self, parser):
        parser.add_argument('--create-csv', dest='create_csv', help='Pasta contendo XMLs para gerar planilha (rascunho)')
        parser.add_argument('--process', action='store_true', dest='process', help='Processar CSV preenchido e gerar XMLs')
        parser.add_argument('--csv', dest='csv', help='Caminho para o CSV preenchido')
        parser.add_argument('--xmls', dest='xmls', help='Pasta que contém os XMLs originais (mesma da etapa 1)')
        parser.add_argument('--output', dest='output', help='Pasta de saída para os XMLs gerados (opcional)')

    def handle(self, *args, **options):
        if options.get('create_csv'):
            pasta = options.get('create_csv')
            try:
                csv_path, count = criar_csv_rascunho(pasta)
                self.stdout.write(self.style.SUCCESS(f'Planilha criada: {csv_path} ({count} notas)'))
            except Exception as e:
                raise CommandError(str(e))

        elif options.get('process'):
            caminho_csv = options.get('csv')
            pasta_xmls = options.get('xmls')
            pasta_saida = options.get('output')
            if not caminho_csv or not pasta_xmls:
                raise CommandError('Para --process é obrigatório fornecer --csv e --xmls')
            try:
                sucesso = processar_final(caminho_csv, pasta_xmls, pasta_saida)
                self.stdout.write(self.style.SUCCESS(f'Processo concluído: {sucesso} documentos gerados.'))
            except Exception as e:
                raise CommandError(str(e))

        else:
            raise CommandError('Use --create-csv ou --process')
