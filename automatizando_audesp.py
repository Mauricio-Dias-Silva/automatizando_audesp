import customtkinter as ctk
from tkinter import filedialog, messagebox, scrolledtext
import xmltodict
import os
import csv
from datetime import datetime

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# --- DADOS FIXOS DO RESPONSÁVEL (Edite aqui uma vez só) ---
RESP_CPF = "00000000000"
RESP_NOME = "NOME DO SECRETARIO"
RESP_CARGO = "SECRETARIO"

class RoboAudespUltimate(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Robô Suprimentos → Audesp (Gerador de CSV + XML)")
        self.geometry("1050x850")
        self.minsize(950, 750)

        # --- FRAME PRINCIPAL COM PADDING ---
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(pady=15, padx=15, fill="both", expand=True)

        # --- HEADER ---
        header_frame = ctk.CTkFrame(main_frame, fg_color="#1f6aa5", corner_radius=10)
        header_frame.pack(fill="x", pady=(0, 20))
        ctk.CTkLabel(header_frame, text="Automação de Notas Fiscais para Audesp", font=("Arial", 18, "bold"), text_color="white").pack(pady=12, padx=15)

        # --- ETAPA 1: SELECIONAR PASTA COM XMLs ---
        etapa1_frame = ctk.CTkFrame(main_frame, fg_color="#fff8e1", corner_radius=8, border_width=2, border_color="#ffc107")
        etapa1_frame.pack(fill="x", pady=(0, 12))
        
        etapa_label = ctk.CTkLabel(etapa1_frame, text="1️⃣  ETAPA 1: Selecionar XMLs de Nota Fiscal", font=("Arial", 13, "bold"), text_color="#f57f17")
        etapa_label.pack(pady=(10, 5), padx=15, anchor="w")
        ctk.CTkLabel(etapa1_frame, text="Clique no botão abaixo para escolher a pasta que contém seus XMLs de nota fiscal", font=("Arial", 9), text_color="#666666").pack(pady=(0, 8), padx=15, anchor="w")
        
        self.btn_gerar_csv = ctk.CTkButton(etapa1_frame, text="📁 Selecionar Pasta com XMLs de Nota Fiscal", command=self.criar_csv_rascunho, height=38, fg_color="#ffc107", text_color="black", font=("Arial", 11, "bold"))
        self.btn_gerar_csv.pack(pady=(0, 10), padx=15, fill="x")
        
        self.status_etapa1 = ctk.CTkLabel(etapa1_frame, text="❌ Aguardando seleção...", font=("Arial", 9), text_color="#d32f2f")
        self.status_etapa1.pack(pady=(0, 10), padx=15, anchor="w")

        # --- ETAPA 1B: SELECIONAR XMLs DE EMPENHOS ---
        etapa1b_frame = ctk.CTkFrame(main_frame, fg_color="#fce4ec", corner_radius=8, border_width=2, border_color="#e91e63")
        etapa1b_frame.pack(fill="x", pady=(0, 12))
        
        etapa_label = ctk.CTkLabel(etapa1b_frame, text="1️⃣ B ETAPA 1B: Selecionar XMLs de Empenhos (Opcional)", font=("Arial", 13, "bold"), text_color="#c2185b")
        etapa_label.pack(pady=(10, 5), padx=15, anchor="w")
        ctk.CTkLabel(etapa1b_frame, text="Se você tem XMLs de empenhos separados, clique para selecioná-los. Caso contrário, deixe em branco.", font=("Arial", 9), text_color="#666666").pack(pady=(0, 8), padx=15, anchor="w")
        
        self.btn_sel_empenhos = ctk.CTkButton(etapa1b_frame, text="📁 Selecionar Pasta com XMLs de Empenhos", command=self.selecionar_xmls_empenhos, height=38, fg_color="#e91e63", text_color="white", font=("Arial", 11, "bold"))
        self.btn_sel_empenhos.pack(pady=(0, 10), padx=15, fill="x")
        
        self.status_etapa1b = ctk.CTkLabel(etapa1b_frame, text="⏭️  Opcional - Clique se tiver XMLs de empenhos", font=("Arial", 9), text_color="#666666")
        self.status_etapa1b.pack(pady=(0, 10), padx=15, anchor="w")

        # --- ETAPA 2: CRIAR CSV ---
        etapa2_frame = ctk.CTkFrame(main_frame, fg_color="#f3e5f5", corner_radius=8, border_width=2, border_color="#9c27b0")
        etapa2_frame.pack(fill="x", pady=(0, 12))
        
        ctk.CTkLabel(etapa2_frame, text="2️⃣  ETAPA 2: Gerar Planilha de Controle", font=("Arial", 13, "bold"), text_color="#6a1b9a").pack(pady=(10, 5), padx=15, anchor="w")
        ctk.CTkLabel(etapa2_frame, text="Um arquivo CSV será gerado automaticamente. Abra-o no Excel e preencha: NUM_EMPENHO, COD_AJUSTE, DATA_EMPENHO e VALOR_TOTAL_EMPENHO", font=("Arial", 9), text_color="#666666").pack(pady=(0, 8), padx=15, anchor="w")
        
        self.status_etapa2 = ctk.CTkLabel(etapa2_frame, text="❌ Aguardando ETAPA 1 ser concluída...", font=("Arial", 9), text_color="#d32f2f")
        self.status_etapa2.pack(pady=(0, 10), padx=15, anchor="w")

        # --- ETAPA 3: SELECIONAR CSV ---
        etapa3_frame = ctk.CTkFrame(main_frame, fg_color="#e8f5e9", corner_radius=8, border_width=2, border_color="#4caf50")
        etapa3_frame.pack(fill="x", pady=(0, 12))
        
        ctk.CTkLabel(etapa3_frame, text="3️⃣  ETAPA 3: Selecionar CSV Preenchido", font=("Arial", 13, "bold"), text_color="#2e7d32").pack(pady=(10, 5), padx=15, anchor="w")
        ctk.CTkLabel(etapa3_frame, text="Após preencher a planilha no Excel, clique para selecioná-la", font=("Arial", 9), text_color="#666666").pack(pady=(0, 8), padx=15, anchor="w")
        
        self.btn_sel_csv = ctk.CTkButton(etapa3_frame, text="📄 Selecionar CSV Preenchido", command=self.selecionar_csv, height=36, fg_color="#4caf50", font=("Arial", 11, "bold"))
        self.btn_sel_csv.pack(pady=(0, 10), padx=15, fill="x")
        
        self.status_etapa3 = ctk.CTkLabel(etapa3_frame, text="❌ Aguardando seleção...", font=("Arial", 9), text_color="#d32f2f")
        self.status_etapa3.pack(pady=(0, 10), padx=15, anchor="w")

        # --- ETAPA 4: CONFIRMAR PASTA DOS XMLs ---
        etapa4_frame = ctk.CTkFrame(main_frame, fg_color="#e1f5fe", corner_radius=8, border_width=2, border_color="#2196f3")
        etapa4_frame.pack(fill="x", pady=(0, 12))
        
        ctk.CTkLabel(etapa4_frame, text="4️⃣  ETAPA 4: Confirmar Pasta dos XMLs Originais", font=("Arial", 13, "bold"), text_color="#01579b").pack(pady=(10, 5), padx=15, anchor="w")
        ctk.CTkLabel(etapa4_frame, text="Selecione a mesma pasta da ETAPA 1 (onde estão os XMLs de nota fiscal)", font=("Arial", 9), text_color="#666666").pack(pady=(0, 8), padx=15, anchor="w")
        
        self.btn_sel_folder = ctk.CTkButton(etapa4_frame, text="📁 Confirmar Pasta dos XMLs", command=self.selecionar_pasta, height=36, fg_color="#2196f3", font=("Arial", 11, "bold"))
        self.btn_sel_folder.pack(pady=(0, 10), padx=15, fill="x")
        
        self.status_etapa4 = ctk.CTkLabel(etapa4_frame, text="❌ Aguardando seleção...", font=("Arial", 9), text_color="#d32f2f")
        self.status_etapa4.pack(pady=(0, 10), padx=15, anchor="w")

        # --- LOG BOX ---
        log_label = ctk.CTkLabel(main_frame, text="📝 Histórico de Operações:", font=("Arial", 10, "bold"), text_color="#333333")
        log_label.pack(pady=(15, 5), anchor="w")

        self.log_box = scrolledtext.ScrolledText(main_frame, height=6, bg="white", fg="black", font=("Courier", 9), relief="flat", bd=1)
        self.log_box.pack(pady=(0, 12), fill="both", expand=True)

        # --- BOTÃO EXECUTAR ---
        self.btn_executar = ctk.CTkButton(main_frame, text="🚀 EXECUTAR E CRIAR XMLs AUDESP", command=self.processar_final, height=45, fg_color="#28a745", font=("Arial", 12, "bold"))
        self.btn_executar.pack(pady=(0, 0), fill="x")

        # Variáveis
        self.caminho_csv = ""
        self.pasta_xmls = ""
        self.pasta_nf_selecionada = ""
        self.pasta_empenhos_selecionada = ""

    def log(self, texto):
        self.log_box.insert("end", texto + "\n")
        self.log_box.see("end")
    
    def atualizar_status(self, etapa, status, mensagem):
        """Atualiza o status visual das etapas"""
        icone = "✅" if status else "❌"
        cor = "#388e3c" if status else "#d32f2f"
        
        if etapa == 1:
            self.status_etapa1.configure(text=f"{icone} {mensagem}", text_color=cor)
        elif etapa == "1b":
            self.status_etapa1b.configure(text=f"{icone} {mensagem}", text_color=cor)
        elif etapa == 2:
            self.status_etapa2.configure(text=f"{icone} {mensagem}", text_color=cor)
        elif etapa == 3:
            self.status_etapa3.configure(text=f"{icone} {mensagem}", text_color=cor)
        elif etapa == 4:
            self.status_etapa4.configure(text=f"{icone} {mensagem}", text_color=cor)

    # --- FUNÇÃO 1B: SELECIONAR XMLs DE EMPENHOS ---
    def selecionar_xmls_empenhos(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta onde estão os XMLs de Empenhos (ou cancele se não tiver)")
        if pasta:
            self.pasta_empenhos_selecionada = pasta
            self.log(f"✓ Pasta de empenhos selecionada: {os.path.basename(pasta)}")
            self.atualizar_status("1b", True, f"Pasta de empenhos selecionada: {os.path.basename(pasta)}")
        else:
            self.log("⏭️  Empenhos ignorados - continuando com notas fiscais")
            self.atualizar_status("1b", True, "Ignorado - usando apenas notas fiscais")

    # --- FUNÇÃO 1: CRIAR O CSV AUTOMÁTICO ---
    def criar_csv_rascunho(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta onde estão as Notas Fiscais (XML)")
        if not pasta: 
            self.atualizar_status(1, False, "Seleção cancelada")
            return

        self.pasta_nf_selecionada = pasta

        arquivos = [f for f in os.listdir(pasta) if f.lower().endswith('.xml')]
        if not arquivos:
            messagebox.showwarning("Vazio", "Nenhum XML encontrado nesta pasta.")
            self.atualizar_status(1, False, "Nenhum XML encontrado")
            return

        # Nome do arquivo CSV de saída
        csv_path = os.path.join(pasta, "PLANILHA_CONTROLE_AUDESP.csv")

        try:
            with open(csv_path, mode='w', newline='', encoding='utf-8-sig') as file:
                writer = csv.writer(file, delimiter=';')
                # Cabeçalho
                writer.writerow(['ARQUIVO_XML', 'FORNECEDOR', 'NUM_NOTA', 'DATA_EMISSAO', 'VALOR_TOTAL_NOTA', 'COD_AJUSTE', 'NUM_EMPENHO', 'DATA_EMPENHO', 'VALOR_TOTAL_EMPENHO', 'VALOR_PARCELA_PAGA'])
                
                count = 0
                for arq in arquivos:
                    try:
                        caminho_full = os.path.join(pasta, arq)
                        with open(caminho_full, "rb") as f: doc = xmltodict.parse(f)
                        
                        # Extração Segura dos dados da NFe
                        inf = doc['nfeProc']['NFe']['infNFe']
                        num = inf['ide']['nNF']
                        data = inf['ide']['dhEmi'].split('T')[0]
                        valor = inf['total']['ICMSTot']['vNF']
                        fornecedor = inf['emit']['xNome']

                        # Escreve a linha. Note que deixamos os campos de Empenho vazios para você preencher
                        # O campo VALOR_PARCELA_PAGA já vem preenchido com o total, você edita se for rateio.
                        writer.writerow([arq, fornecedor, num, data, valor, '', '', '', '', valor])
                        count += 1
                    except Exception as e:
                        self.log(f"Erro ao ler {arq}: {e}")

            mensagem_sucesso = f"Planilha criada com {count} notas! Local: {csv_path}"
            messagebox.showinfo("Sucesso", f"Planilha criada!\nLocal: {csv_path}\n\nAbra este arquivo no Excel e preencha:\n• NUM_EMPENHO\n• COD_AJUSTE\n• DATA_EMPENHO\n• VALOR_TOTAL_EMPENHO")
            self.log(f"✓ ETAPA 1 CONCLUÍDA: Planilha de controle gerada com {count} notas!")
            self.atualizar_status(1, True, f"Pasta selecionada - {count} XMLs encontrados")
            self.atualizar_status(2, True, "Planilha criada! Agora preencha no Excel...")
            
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.atualizar_status(1, False, f"Erro: {str(e)[:40]}")

    # --- FUNÇÃO 2: PROCESSAR E GERAR AUDESP ---
    def selecionar_csv(self):
        self.caminho_csv = filedialog.askopenfilename(title="Selecione o arquivo CSV que você preencheu", filetypes=[("CSV Files", "*.csv")])
        if self.caminho_csv:
            self.log(f"✓ CSV selecionado: {os.path.basename(self.caminho_csv)}")
            self.atualizar_status(3, True, f"CSV selecionado: {os.path.basename(self.caminho_csv)}")
        else:
            self.atualizar_status(3, False, "Seleção cancelada")

    def selecionar_pasta(self):
        self.pasta_xmls = filedialog.askdirectory(title="Selecione a pasta que contém os XMLs originais (mesma da ETAPA 1)")
        if self.pasta_xmls:
            self.log(f"✓ Pasta confirmada: {os.path.basename(self.pasta_xmls)}")
            self.atualizar_status(4, True, f"Pasta confirmada: {os.path.basename(self.pasta_xmls)}")
        else:
            self.atualizar_status(4, False, "Seleção cancelada")

    def processar_final(self):
        if not self.caminho_csv or not self.pasta_xmls:
            messagebox.showwarning("Atenção", "Selecione o CSV preenchido e a Pasta dos XMLs.")
            return
        
        pasta_saida = os.path.join(self.pasta_xmls, "SAIDA_AUDESP_PRONTOS")
        if not os.path.exists(pasta_saida): os.makedirs(pasta_saida)

        sucesso = 0
        try:
            with open(self.caminho_csv, newline='', encoding='utf-8-sig') as csvfile:
                reader = csv.DictReader(csvfile, delimiter=';')
                
                for row in reader:
                    # Pula linhas vazias se houver
                    if not row['COD_AJUSTE'] or not row['NUM_EMPENHO']:
                        self.log(f"PULADO: Nota {row.get('NUM_NOTA')} sem dados de empenho.")
                        continue

                    # Dados do CSV
                    arq_xml = row['ARQUIVO_XML']
                    ajuste = row['COD_AJUSTE']
                    num_empenho = row['NUM_EMPENHO']
                    data_empenho = row['DATA_EMPENHO']
                    
                    try:
                        # Tratamento de número (Excel brasileiro usa vírgula, Python usa ponto)
                        val_total_nota = float(row['VALOR_TOTAL_NOTA'].replace(',', '.'))
                        val_pago = float(row['VALOR_PARCELA_PAGA'].replace(',', '.'))
                        val_total_empenho = float(row['VALOR_TOTAL_EMPENHO'].replace(',', '.')) if row['VALOR_TOTAL_EMPENHO'] else val_total_nota
                        
                        # Calcula percentual
                        percentual = (val_pago / val_total_empenho * 100) if val_total_empenho > 0 else 0
                        
                        # --- GERAÇÃO DOS ARQUIVOS ---
                        self.gerar_xml_fisico(pasta_saida, row, val_pago, val_total_empenho, percentual)
                        sucesso += 1
                        
                    except ValueError:
                        self.log(f"Erro de valor na nota {row.get('NUM_NOTA')}")

            messagebox.showinfo("Fim", f"Processo concluído! {sucesso} documentos gerados na pasta SAIDA_AUDESP_PRONTOS.")

        except Exception as e:
            messagebox.showerror("Erro Fatal", str(e))

    def gerar_xml_fisico(self, pasta, dados, valor_pago, valor_total_empenho, percentual):
        """Gera os 3 arquivos XML necessários para o Audesp"""
        
        num_nota = dados['NUM_NOTA']
        num_empenho = dados['NUM_EMPENHO']
        cod_ajuste = dados['COD_AJUSTE']
        data_empenho = dados['DATA_EMPENHO']
        fornecedor = dados['FORNECEDOR']
        
        prefixo = f"{num_nota}_EMP{num_empenho}"
        
        try:
            # --- 1. XML DE EXECUÇÃO ---
            xml_exec = f"""<?xml version="1.0" encoding="UTF-8"?>
<execucao>
    <numero_empenho>{num_empenho}</numero_empenho>
    <data_empenho>{data_empenho}</data_empenho>
    <codigo_ajuste>{cod_ajuste}</codigo_ajuste>
    <valor_total>{valor_total_empenho:.2f}</valor_total>
    <fornecedor>{fornecedor}</fornecedor>
    <data_criacao>{datetime.now().strftime('%Y-%m-%d')}</data_criacao>
</execucao>"""
            
            with open(os.path.join(pasta, f"{prefixo}_EXECUCAO.xml"), 'w', encoding='utf-8') as f:
                f.write(xml_exec)
            
            # --- 2. XML FISCAL ---
            xml_fiscal = f"""<?xml version="1.0" encoding="UTF-8"?>
<fiscal>
    <numero_nota>{num_nota}</numero_nota>
    <numero_empenho>{num_empenho}</numero_empenho>
    <valor_nota>{valor_total_empenho:.2f}</valor_nota>
    <valor_fiscalizado>{valor_pago:.2f}</valor_fiscalizado>
    <percentual_pago>{percentual:.2f}</percentual_pago>
    <responsavel_cpf>{RESP_CPF}</responsavel_cpf>
    <responsavel_nome>{RESP_NOME}</responsavel_nome>
    <responsavel_cargo>{RESP_CARGO}</responsavel_cargo>
    <data_fiscalizacao>{datetime.now().strftime('%Y-%m-%d')}</data_fiscalizacao>
</fiscal>"""
            
            with open(os.path.join(pasta, f"{prefixo}_FISCAL.xml"), 'w', encoding='utf-8') as f:
                f.write(xml_fiscal)
            
            # --- 3. XML DE PAGAMENTO ---
            xml_pago = f"""<?xml version="1.0" encoding="UTF-8"?>
<pagamento>
    <numero_empenho>{num_empenho}</numero_empenho>
    <data_pagamento>{datetime.now().strftime('%Y-%m-%d')}</data_pagamento>
    <valor_pago>{valor_pago:.2f}</valor_pago>
    <percentual_pago>{percentual:.2f}</percentual_pago>
    <fornecedor>{fornecedor}</fornecedor>
    <status>PAGO</status>
</pagamento>"""
            
            with open(os.path.join(pasta, f"{prefixo}_PAGAMENTO.xml"), 'w', encoding='utf-8') as f:
                f.write(xml_pago)
            
            self.log(f"✓ Gerados 3 XMLs para Nota {num_nota} (Empenho {num_empenho})")
            
        except Exception as e:
            self.log(f"✗ Erro ao gerar XMLs para {num_nota}: {e}")


if __name__ == "__main__":
    app = RoboAudespUltimate()
    app.mainloop()