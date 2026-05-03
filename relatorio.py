import io, os
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

STATIC = Path(__file__).parent

AZUL       = "1A5276"
VERMELHO   = "FF0000"
VERM_ESC   = "8E1A0E"
BRANCO     = "FFFFFF"
CINZA      = "F2F2F2"
ROSA       = "FFF0F0"

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color=BRANCO, size=10):
    return Font(bold=bold, color=color, size=size)

def _align(h="center"):
    return Alignment(horizontal=h, vertical="center", wrap_text=True)

# ── Excel ─────────────────────────────────────────────────────
def gerar_excel_resultado(resultado, nome_busca, cpf_busca):
    wb = Workbook()
    wb.remove(wb.active)
    anos = sorted(resultado.keys())

    # Aba RESUMO
    ws = wb.create_sheet("RESUMO", 0)
    ws.merge_cells("A1:D1")
    ws["A1"] = "AGÊNCIA DE DEFESA AGROPECUÁRIA DO ESTADO DO PARÁ"
    ws["A1"].font      = Font(bold=True, size=13, color=AZUL)
    ws["A1"].alignment = _align()

    ws["A3"] = "Produtor pesquisado:"
    ws["B3"] = (nome_busca or '').upper()
    ws["A4"] = "CPF/CNPJ:"
    ws["B4"] = cpf_busca or ''
    ws["A5"] = "Data do relatório:"
    ws["B5"] = datetime.now().strftime("%d/%m/%Y %H:%M")

    for col, titulo in enumerate(["ANO","GTAs como REMETENTE","GTAs como DESTINATÁRIO","TOTAL"], 1):
        c = ws.cell(row=7, column=col, value=titulo)
        c.fill = _fill(AZUL); c.font = _font(bold=True); c.alignment = _align()

    tot_o = tot_d = 0
    for i, ano in enumerate(anos):
        qo = len(resultado[ano]['origem'])
        qd = len(resultado[ano]['destino'])
        ws.cell(row=8+i, column=1, value=str(ano))
        ws.cell(row=8+i, column=2, value=qo)
        ws.cell(row=8+i, column=3, value=qd)
        ws.cell(row=8+i, column=4, value=qo+qd)
        tot_o += qo; tot_d += qd

    lr = 8 + len(anos)
    for col, val in enumerate(["TOTAL", tot_o, tot_d, tot_o+tot_d], 1):
        c = ws.cell(row=lr, column=col, value=val)
        c.font = Font(bold=True)

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 25
    ws.column_dimensions["D"].width = 12

    # Uma aba por ano
    COLS_INTERNAS = {'_ano','_arquivo','_formato'}
    for ano in anos:
        orig = resultado[ano]['origem']
        dest = resultado[ano]['destino']
        cols = resultado[ano]['colunas']
        cols = [c for c in cols if c not in COLS_INTERNAS]
        if not orig and not dest:
            continue

        ws_ano = wb.create_sheet(str(ano))
        linha  = 1

        def escrever_bloco(dados, fill_header, fill_zebra, titulo_sep=None):
            nonlocal linha
            if not dados:
                return
            for ci, col in enumerate(cols, 1):
                c = ws_ano.cell(row=linha, column=ci, value=col)
                c.fill = _fill(fill_header)
                c.font = _font(bold=True)
                c.alignment = _align()
            ws_ano.row_dimensions[linha].height = 28
            linha += 1
            for ri, row_dict in enumerate(dados):
                for ci, col in enumerate(cols, 1):
                    val = row_dict.get(col, '')
                    ws_ano.cell(row=linha, column=ci, value=val)
                    if ri % 2 == 0:
                        ws_ano.cell(row=linha, column=ci).fill = _fill(fill_zebra)
                linha += 1

        escrever_bloco(orig, AZUL, CINZA)

        for ci in range(1, len(cols)+1):
            c = ws_ano.cell(row=linha, column=ci)
            c.fill = _fill(VERMELHO)
            c.font = _font(bold=True)
            c.alignment = _align()
        ws_ano.cell(row=linha, column=1).value = "▼  DESTINATÁRIO  ▼"
        ws_ano.row_dimensions[linha].height = 18
        linha += 1

        escrever_bloco(dest, VERM_ESC, "FFF0F0")

        for ci, col in enumerate(cols, 1):
            letra = get_column_letter(ci)
            tam = min(max(len(str(col)), 10) + 2, 40)
            ws_ano.column_dimensions[letra].width = tam

        ws_ano.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── PDF de Auditoria ──────────────────────────────────────────
def gerar_pdf_auditoria(rows, gerado_por, data_ini='', data_fim=''):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=4*cm, bottomMargin=1.5*cm
    )

    styles  = getSampleStyleSheet()
    cor_azul = colors.HexColor('#1A5276')

    elementos = []

    periodo = f"{data_ini or 'início'} a {data_fim or 'hoje'}"
    elementos.append(Paragraph(
        "<b>Relatório de Auditoria de Acessos</b>",
        ParagraphStyle('titulo', fontSize=13, textColor=cor_azul,
                       alignment=TA_CENTER, spaceAfter=4)
    ))
    elementos.append(Paragraph(
        f"Período: {periodo} &nbsp;&nbsp; | &nbsp;&nbsp; Gerado por: {gerado_por.get('nome','')} &nbsp;&nbsp; | &nbsp;&nbsp; {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        ParagraphStyle('sub', fontSize=9, textColor=colors.grey,
                       alignment=TA_CENTER, spaceAfter=12)
    ))

    cabecalho = [
        "Data/Hora", "Usuário", "Órgão", "IP",
        "Localidade", "CPF/CNPJ\npesquisado", "Nome\npesquisado"
    ]
    dados_tabela = [cabecalho]
    for r in rows:
        dados_tabela.append([
            r.get('data_hora',''),
            r.get('usuario_nome',''),
            r.get('orgao',''),
            r.get('ip',''),
            r.get('localidade',''),
            r.get('cpf_cnpj_pesquisado',''),
            r.get('nome_pesquisado',''),
        ])

    col_widths = [2.8*cm, 4*cm, 3*cm, 2.5*cm, 2.5*cm, 3*cm, 4.2*cm]

    tabela = Table(dados_tabela, colWidths=col_widths, repeatRows=1)
    tabela.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,0), cor_azul),
        ('TEXTCOLOR',     (0,0), (-1,0), colors.white),
        ('FONTNAME',      (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',      (0,0), (-1,0), 8),
        ('ALIGN',         (0,0), (-1,0), 'CENTER'),
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ('FONTSIZE',      (0,1), (-1,-1), 7),
        ('ROWBACKGROUNDS',(0,1), (-1,-1), [colors.white, colors.HexColor('#F2F2F2')]),
        ('GRID',          (0,0), (-1,-1), 0.3, colors.HexColor('#CCCCCC')),
        ('TOPPADDING',    (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elementos.append(tabela)

    brasao_path  = str(STATIC / 'brasao_para.png')
    adepara_path = str(STATIC / 'adepara_logo.png')

    def header(canvas, doc):
        canvas.saveState()
        w, h = A4

        if os.path.exists(brasao_path):
            canvas.drawImage(brasao_path, 1.5*cm, h-3.5*cm, width=2*cm, height=2.5*cm,
                             preserveAspectRatio=True, mask='auto')

        if os.path.exists(adepara_path):
            canvas.drawImage(adepara_path, w-4.5*cm, h-3.5*cm, width=3*cm, height=2*cm,
                             preserveAspectRatio=True, mask='auto')

        canvas.setFont("Helvetica-Bold", 10)
        canvas.setFillColor(cor_azul)
        canvas.drawCentredString(w/2, h-1.5*cm, "AGÊNCIA DE DEFESA AGROPECUÁRIA DO ESTADO DO PARÁ")

        canvas.setStrokeColor(cor_azul)
        canvas.setLineWidth(0.5)
        canvas.line(1.5*cm, h-3.7*cm, w-1.5*cm, h-3.7*cm)

        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.grey)
        canvas.drawCentredString(w/2, 0.8*cm,
            f"Página {doc.page} — Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')} — Sistema de Consulta GTA")
        canvas.restoreState()

    doc.build(elementos, onFirstPage=header, onLaterPages=header)
    buf.seek(0)
    return buf
