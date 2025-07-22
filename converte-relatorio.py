import os
import shutil
from datetime import datetime, timedelta
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.datetime import from_excel
from copy import copy
import xlrd
from pathlib import Path

pasta_projeto = str(Path().resolve())
nome_base = 'Relatorio-Base.xlsx'
caminho_base = os.path.join(pasta_projeto, nome_base)

# Busca arquivos Excel na pasta do projeto
arquivos_excel = [
    f for f in os.listdir(pasta_projeto)
    if (f.endswith('.xls') or f.endswith('.xlsx')) and not f.startswith("Relatorio-Base")
]
if not arquivos_excel:
    print("[ERRO] Nenhum arquivo Excel (.xls ou .xlsx) encontrado na pasta do projeto.")
    exit()

arquivos_excel_path = [os.path.join(pasta_projeto, f) for f in arquivos_excel]
arquivo_novo = max(arquivos_excel_path, key=os.path.getctime)

def encontrar_proxima_linha_vazia(ws, min_row=3):
    for i in range(min_row, ws.max_row + 2):
        # Para cada linha i, verifica se TODAS as células na linha são None ou strings vazias (ou só espaços)
        if all((cell.value is None) or (str(cell.value).strip() == '') for cell in ws[i]):
            return i
    return ws.max_row + 1


def formata_arquivo_novo(wb):
    ws = wb.active

    verde_escuro = "006400"
    fonte_cabecalho = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
    fill_cabecalho = PatternFill(start_color=verde_escuro, end_color=verde_escuro, fill_type="solid")
    alinhamento_cabecalho = Alignment(horizontal="center", vertical="center", wrap_text=True)
    borda_fina = Border(
        left=Side(style='thin', color='FFFFFF'),
        right=Side(style='thin', color='FFFFFF'),
        top=Side(style='thin', color='FFFFFF'),
        bottom=Side(style='thin', color='FFFFFF')
    )

    # Mesclar células
    ws.merge_cells('A1:B1')
    ws.merge_cells('C1:S1')
    ws.merge_cells('T1:T2')

    ws['A1'] = "CONTA"
    ws['C1'] = "PROTOCOLO"
    ws['T1'] = "MOVIMENTAÇÕES"

    conta_cabecalho = ["IDENTIFICADOR", "NOME DA CONTA"]
    protocolo_cabecalho = [
        "DATA DE ABERTURA", "PROTOCOLO", "STATUS", "STATUS GPU", "RESPONSAVEL", "RESPONSAVEL ORIGEM",
        "DATA DE ENCERRAMENTO", "DESCRIÇÃO DE ATENDIMENTO", "DESCRIÇÃO DE ENCERRAMENTO", "DESCRIÇÃO",
        "CATEGORIA", "MOTIVO", "SUBMOTIVO", "CANAL DE ENTRADA", "GRUPO ATENDIMENTO ATUAL",
        "GRUPO ATENDIMENTO ORIGEM", "TIPO DE SOLICITAÇÃO"
    ]

    for i, titulo in enumerate(conta_cabecalho, start=1):
        cell = ws.cell(row=2, column=i, value=titulo)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = alinhamento_cabecalho
        cell.border = borda_fina

    for i, titulo in enumerate(protocolo_cabecalho, start=3):
        cell = ws.cell(row=2, column=i, value=titulo)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = alinhamento_cabecalho
        cell.border = borda_fina
        if i in (3, 7):
            cell.number_format = "DD/MM/YYYY"

    ws['T1'].font = fonte_cabecalho
    ws['T1'].fill = fill_cabecalho
    ws['T1'].alignment = alinhamento_cabecalho
    ws['T1'].border = borda_fina

    for col in range(1, 21):
        for row in [1, 2]:
            cell = ws.cell(row=row, column=col)
            cell.font = fonte_cabecalho
            cell.fill = fill_cabecalho
            cell.alignment = alinhamento_cabecalho
            cell.border = borda_fina

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 30

    larguras = [20, 25, 15, 30, 20, 20, 20, 20, 25, 20, 25, 20, 20, 25, 20, 25, 20, 25, 30, 40]
    for i, largura in enumerate(larguras, start=1):
        letra_col = chr(64 + i)
        ws.column_dimensions[letra_col].width = largura


def converter_xls_para_xlsx(caminho_xls, caminho_xlsx):
    from openpyxl.utils.datetime import from_excel

    wb_xls = xlrd.open_workbook(caminho_xls)
    sheet_xls = wb_xls.sheet_by_index(0)

    wb_xlsx = Workbook()
    ws_xlsx = wb_xlsx.active

    altura_linha_fina = 15
    colunas_datas = [3, 7]  # colunas 1-based que sao datas

    for row_idx in range(sheet_xls.nrows):
        for col_idx in range(sheet_xls.ncols):
            valor = sheet_xls.cell_value(row_idx, col_idx)
            cell = ws_xlsx.cell(row=row_idx + 1, column=col_idx + 1)

            # Se for coluna de data e valor for número, tenta converter
            if (col_idx + 1) in colunas_datas and isinstance(valor, (int, float)):
                try:
                    cell.value = from_excel(valor)
                    cell.number_format = "DD/MM/YYYY"
                except:
                    cell.value = valor
            else:
                cell.value = valor

        # Força altura da linha fininha
        ws_xlsx.row_dimensions[row_idx + 1].height = altura_linha_fina

    wb_xlsx.save(caminho_xlsx)
    print(f"[OK] Convertido .xls para .xlsx com data corrigida e altura padronizada: {caminho_xlsx}")

def init_xlsx(caminho_arquivo):
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório"

    # Cores e fontes
    verde_escuro = "006400"  # dark green
    fonte_cabecalho = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
    fill_cabecalho = PatternFill(start_color=verde_escuro, end_color=verde_escuro, fill_type="solid")
    alinhamento_cabecalho = Alignment(horizontal="center", vertical="center", wrap_text=True)
    borda_fina = Border(
        left=Side(style='thin', color='FFFFFF'),
        right=Side(style='thin', color='FFFFFF'),
        top=Side(style='thin', color='FFFFFF'),
        bottom=Side(style='thin', color='FFFFFF')
    )

    # Mesclar células para o cabeçalho linha 1
    ws.merge_cells('A1:B1')  # CONTA
    ws.merge_cells('C1:S1')  # PROTOCOLO (colunas C a S)
    ws.merge_cells('T1:T2')  # MOVIMENTAÇÕES (coluna T)

    # Texto cabeçalho linha 1
    ws['A1'] = "CONTA"
    ws['C1'] = "PROTOCOLO"
    ws['T1'] = "MOVIMENTAÇÕES"

    # Cabeçalho linha 2 (detalhado)
    conta_cabecalho = ["IDENTIFICADOR", "NOME DA CONTA"]
    protocolo_cabecalho = [
        "DATA DE ABERTURA", "PROTOCOLO", "STATUS", "STATUS GPU", "RESPONSAVEL", "RESPONSAVEL ORIGEM",
        "DATA DE ENCERRAMENTO", "DESCRIÇÃO DE ATENDIMENTO", "DESCRIÇÃO DE ENCERRAMENTO", "DESCRIÇÃO",
        "CATEGORIA", "MOTIVO", "SUBMOTIVO", "CANAL DE ENTRADA", "GRUPO ATENDIMENTO ATUAL",
        "GRUPO ATENDIMENTO ORIGEM", "TIPO DE SOLICITAÇÃO"
    ]

    # Preencher conta (A2, B2)
    for i, titulo in enumerate(conta_cabecalho, start=1):
        cell = ws.cell(row=2, column=i, value=titulo)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = alinhamento_cabecalho
        cell.border = borda_fina

    # Preencher protocolo (C2 até S2, 17 colunas)
    for i, titulo in enumerate(protocolo_cabecalho, start=3):
        cell = ws.cell(row=2, column=i, value=titulo)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = alinhamento_cabecalho
        cell.border = borda_fina
        # Se for colunas de data, já aplica formatação aqui para segurança
        if i in (3, 7):
            cell.number_format = "DD/MM/YYYY"

    # Preencher MOVIMENTAÇÕES (T1 e T2 já mescladas)
    cell = ws['T1']
    cell.font = fonte_cabecalho
    cell.fill = fill_cabecalho
    cell.alignment = alinhamento_cabecalho
    cell.border = borda_fina

    # Estilo das células mescladas da linha 1 e 2
    for col in range(1, 21):
        for row in [1, 2]:
            cell = ws.cell(row=row, column=col)
            if (row == 1 and (col <= 2 or (3 <= col <= 19) or col == 20)) or row == 2:
                cell.font = fonte_cabecalho
                cell.fill = fill_cabecalho
                cell.alignment = alinhamento_cabecalho
                cell.border = borda_fina

    # Ajustar altura das linhas do cabeçalho
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 30

    # Ajustar largura colunas
    larguras = [20, 25, 15, 30, 20, 20, 20, 20, 25, 20, 25, 20, 20, 25, 20, 25, 20, 25, 30, 40]
    for i, largura in enumerate(larguras, start=1):
        letra_col = chr(64 + i)
        ws.column_dimensions[letra_col].width = largura

    wb.save(caminho_arquivo)
    print(f"[OK] Planilha base inicializada e salva em: {caminho_arquivo}")


# Se o arquivo for .xls, converte para .xlsx antes de seguir
if arquivo_novo.lower().endswith('.xls'):
    caminho_convertido = os.path.splitext(arquivo_novo)[0] + '.xlsx'  # 'Relatorio.xls' -> 'Relatorio.xlsx'
    converter_xls_para_xlsx(arquivo_novo, caminho_convertido)
    caminho_para_uso = caminho_convertido
else:
    caminho_para_uso = arquivo_novo

data_ontem = (datetime.now() - timedelta(days=4)).strftime('%Y-%m-%d')

# Colunas que são datas no relatório (1-based)
colunas_datas = [3, 7]

altura_linha_fina = 15

if not os.path.exists(caminho_base):
    # Cenário 1: base não existe → cria base formatada e copia dados
    init_xlsx(caminho_base)
    wb_base = load_workbook(caminho_base)
    ws_base = wb_base.active

    wb_novo = load_workbook(caminho_para_uso)
    ws_novo = wb_novo.active

    # Copia dados para o base
    for i, row in enumerate(ws_novo.iter_rows(min_row=3, values_only=True), start=3):
        for col_idx, valor in enumerate(row, start=1):
            if col_idx in colunas_datas and isinstance(valor, (int, float)):
                try: valor = from_excel(valor)
                except: pass
            ws_base.cell(row=i, column=col_idx, value=valor)
            if col_idx in colunas_datas:
                ws_base.cell(row=i, column=col_idx).number_format = "DD/MM/YYYY"
        ws_base.row_dimensions[i].height = altura_linha_fina

    wb_base.save(caminho_base)
    print(f"[OK] Arquivo base criado com dados do arquivo convertido: {caminho_base}")

    # → Agora gera também o arquivo diário formatado igual à base
    nome_diario = f"Relatorio-{data_ontem}.xlsx"
    caminho_diario = os.path.join(pasta_projeto, nome_diario)
    wb_diario = Workbook()
    formata_arquivo_novo(wb_diario)
    ws_diario = wb_diario.active

    # Copia os dados do base para o diário
    for i, row in enumerate(ws_base.iter_rows(min_row=3, values_only=True), start=3):
        for col_idx, valor in enumerate(row, start=1):
            cell = ws_diario.cell(row=i, column=col_idx, value=valor)
            if col_idx in colunas_datas:
                cell.number_format = "DD/MM/YYYY"
        ws_diario.row_dimensions[i].height = altura_linha_fina

    wb_diario.save(caminho_diario)
    print(f"[OK] Arquivo diário criado e formatado: {caminho_diario}")

else:
    # Cenário 2: base existe → move e formata o diário antes de colar na base
    nome_diario = f"Relatorio-{data_ontem}.xlsx"
    caminho_diario = os.path.join(pasta_projeto, nome_diario)
    shutil.move(caminho_para_uso, caminho_diario)
    print(f"[OK] Novo relatório movido para: {caminho_diario}")

    # Aplica formatação igual à base
    wb_diario = load_workbook(caminho_diario)
    formata_arquivo_novo(wb_diario)
    wb_diario.save(caminho_diario)

    # Agora cola os dados formatados no base
    wb_base = load_workbook(caminho_base); ws_base = wb_base.active
    ws_diario = wb_diario.active

    linha_modelo = ws_base.max_row
    proxima = encontrar_proxima_linha_vazia(ws_base, min_row=3)

    for i, row in enumerate(ws_diario.iter_rows(min_row=3), start=proxima):
        for col_idx, cel in enumerate(row, start=1):
            valor = cel.value
            if col_idx in colunas_datas and isinstance(valor, (int, float)):
                try: valor = from_excel(valor)
                except: pass
            dest_cell = ws_base.cell(row=i, column=col_idx, value=valor)
            if col_idx in colunas_datas:
                dest_cell.number_format = "DD/MM/YYYY"
            # copia estilo do modelo
            modelo_cel = ws_base.cell(row=linha_modelo, column=col_idx)
            dest_cell.font   = copy(modelo_cel.font)
            dest_cell.fill   = copy(modelo_cel.fill)
            dest_cell.border = copy(modelo_cel.border)
            dest_cell.alignment   = copy(modelo_cel.alignment)
            dest_cell.protection  = copy(modelo_cel.protection)
        altura_modelo = ws_base.row_dimensions[linha_modelo].height or altura_linha_fina
        ws_base.row_dimensions[i].height = altura_modelo

    wb_base.save(caminho_base)
    print(f"[OK] Dados do relatório diário colados no base, formatação mantida.")

# --- LIMPEZA FINAL: remove todos os .xls da pasta do projeto ---
for nome_arquivo in os.listdir(pasta_projeto):
    if nome_arquivo.lower().endswith('.xls'):
        caminho_arquivo = os.path.join(pasta_projeto, nome_arquivo)
        try:
            os.remove(caminho_arquivo)
            print(f"[OK] Removido arquivo temporário: {nome_arquivo}")
        except Exception as e:
            print(f"[ERRO] Falha ao remover {nome_arquivo}: {e}")
