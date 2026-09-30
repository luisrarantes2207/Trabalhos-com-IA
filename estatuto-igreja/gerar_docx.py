"""Gera os .docx (estatuto, ata, edital e requerimento) a partir dos arquivos .txt.

Os artigos do estatuto são numerados automaticamente, e as remissões [[chave]]
(em qualquer dos arquivos) apontam para a numeração do estatuto.
"""
import re, sys
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

base = Path(__file__).parent
DOCS = [("estatuto.txt", "Estatuto_Social_Igreja_ISAC_v2.docx"),
        ("ata_fundacao.txt", "Ata_Assembleia_Fundacao_Igreja_ISAC.docx"),
        ("edital_convocacao.txt", "Edital_Convocacao_Fundacao_Igreja_ISAC.docx"),
        ("requerimento_registro.txt", "Requerimento_Registro_RCPJ_Igreja_ISAC.docx"),
        ("regimento_interno.txt", "Regimento_Interno_Igreja_ISAC.docx"),
        ("codigo_conduta.txt", "Codigo_de_Conduta_Igreja_ISAC.docx"),
        ("politica_protecao.txt", "Politica_Protecao_Menores_Vulneraveis_Igreja_ISAC.docx"),
        ("politica_privacidade.txt", "Politica_Privacidade_Protecao_Dados_Igreja_ISAC.docx"),
        ("termo_voluntariado.txt", "Termo_Adesao_Servico_Voluntario_Igreja_ISAC.docx"),
        ("contrato_comodato.txt", "Contrato_Comodato_Templo_Igreja_ISAC.docx"),
        ("contrato_locacao.txt", "Contrato_Locacao_Templo_Igreja_ISAC.docx"),
        ("recibo_dizimos.txt", "Recibo_Dizimos_Ofertas_Igreja_ISAC.docx"),
        ("termo_doacao.txt", "Termo_Doacao_Bens_Igreja_ISAC.docx"),
        ("termo_doacao_imovel.txt", "Termo_Doacao_Imovel_Igreja_ISAC.docx"),
        ("contrato_trabalho.txt", "Contrato_Trabalho_Empregado_Igreja_ISAC.docx"),
        ("contrato_servicos.txt", "Contrato_Prestacao_Servicos_Igreja_ISAC.docx"),
        ("rpa.txt", "RPA_Recibo_Pagamento_Autonomo_Igreja_ISAC.docx"),
        ("ata_diretoria.txt", "Ata_Reuniao_Diretoria_Executiva_Igreja_ISAC.docx")]

def label(n):
    return f"{n}º" if n < 10 else f"{n}"

# Numeração a partir do estatuto
nums, n = {}, 0
for line in (base / "estatuto.txt").read_text(encoding="utf-8").splitlines():
    m = re.match(r"ART\[(\w+)\]", line)
    if m:
        n += 1
        nums[m.group(1)] = n

# Numeração do Regimento Interno, para remissões <<chave>> de outros documentos
ri_nums, n = {}, 0
for line in (base / "regimento_interno.txt").read_text(encoding="utf-8").splitlines():
    m = re.match(r"RART\[(\w+)\]", line)
    if m:
        n += 1
        ri_nums[m.group(1)] = n

def doc_nums(stem):
    table, k = {}, 0
    for line in (base / f"{stem}.txt").read_text(encoding="utf-8").splitlines():
        m = re.match(r"RART\[(\w+)\]", line)
        if m:
            k += 1
            table[m.group(1)] = k
    return table

def refs(text, local=None, clauses=None):
    """[[chave]] remete ao Estatuto, <<chave>> ao Regimento Interno, {{chave}} a artigo (RART)
    e ((chave)) a cláusula (CL) do próprio documento."""
    def clause(m):
        if m.group(1) not in (clauses or {}):
            sys.exit(f"Referência inexistente: {m.group(0)}")
        return f"{clauses[m.group(1)]}ª"
    text = re.sub(r"\(\((\w+)\)\)", clause, text)
    def sub(table, m):
        k = m.group(1)
        if k not in table:
            sys.exit(f"Referência inexistente: {k}")
        return label(table[k])
    text = re.sub(r"\[\[(\w+)\]\]", lambda m: sub(nums, m), text)
    text = re.sub(r"<<(\w+)>>", lambda m: sub(ri_nums, m), text)
    # {{arquivo:chave}} remete a artigo de outro documento (ex.: {{codigo_conduta:cons}})
    text = re.sub(r"\{\{(\w+):(\w+)\}\}",
                  lambda m: label(doc_nums(m.group(1)).get(m.group(2)) or sys.exit(f"Referência inexistente: {m.group(0)}")),
                  text)
    return re.sub(r"\{\{(\w+)\}\}", lambda m: sub(local or {}, m), text)

def build(src_name, out_name):
    src = (base / src_name).read_text(encoding="utf-8").splitlines()
    local, k = {}, 0
    for line in src:
        m = re.match(r"RART\[(\w+)\]", line)
        if m:
            k += 1
            local[m.group(1)] = k
    clauses = {}
    for line in src:
        m = re.match(r"CL\[(\w+)\]", line)
        if m:
            clauses[m.group(1)] = len(clauses) + 1
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21)
    sec.top_margin = sec.left_margin = Cm(3)
    sec.bottom_margin = sec.right_margin = Cm(2)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Times New Roman", Pt(12)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.line_spacing = 1.15

    def para(text="", bold_prefix=None, align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=False, size=None, indent=None, italic=False):
        p = doc.add_paragraph()
        p.alignment = align
        if indent is not None:
            p.paragraph_format.left_indent = Cm(indent)
        if bold_prefix:
            r = p.add_run(bold_prefix); r.bold = True
        r = p.add_run(text); r.bold = bold; r.italic = italic
        if size:
            for r in p.runs: r.font.size = Pt(size)
        return p

    def date_line():
        p = para("[CIDADE]/[UF], [DIA] de [MÊS] de [ANO].", align=WD_ALIGN_PARAGRAPH.RIGHT)
        p.paragraph_format.space_before = Pt(18)

    def signatures(pairs):
        for nome, cargo in pairs:
            p = para("__________________________________________", align=WD_ALIGN_PARAGRAPH.CENTER)
            p.paragraph_format.space_before = Pt(30); p.paragraph_format.space_after = Pt(0)
            if nome:
                p = para(nome, align=WD_ALIGN_PARAGRAPH.CENTER); p.paragraph_format.space_after = Pt(0)
            para(cargo, align=WD_ALIGN_PARAGRAPH.CENTER)

    grid = None  # linhas acumuladas de uma tabela preenchida (%GRID ... %ENDGRID)

    def flush_grid(rows, widths=None):
        cols = len(rows[0])
        t = doc.add_table(rows=len(rows), cols=cols)
        t.style = "Table Grid"
        for i, row in enumerate(rows):
            for j in range(cols):
                cell = t.rows[i].cells[j]
                cell.text = ""
                r = cell.paragraphs[0].add_run(row[j] if j < len(row) else "")
                r.font.size = Pt(9.5); r.bold = (i == 0)
                cell.width = Cm(widths[j] if widths else 16 / cols)
        doc.add_paragraph()

    for line in src:
        line = refs(line.rstrip(), local, clauses)
        if grid is not None:
            if line.startswith("%ENDGRID"):
                flush_grid(grid, grid_widths); grid = None
            elif line:
                grid.append([c.strip() for c in line.split("|")])
            continue
        if line.startswith("%GRID"):
            grid = []
            # larguras opcionais em cm: "%GRID 5,11"
            spec = line[5:].strip()
            grid_widths = [float(w) for w in spec.split(",")] if spec else None
            continue
        if not line:
            continue
        if line.startswith("# "):
            para(line[2:], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14)
        elif line.startswith("## "):
            p = para(line[3:], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
            p.paragraph_format.space_before = Pt(12)
        elif line.startswith("### "):
            para(line[4:], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, italic=True)
        elif line.startswith("!"):
            para(line[1:], align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)
        elif line.startswith(("ART[", "RART[")):
            k = re.match(r"R?ART\[(\w+)\]\s*", line)
            n = (local if line.startswith("R") else nums)[k.group(1)]
            para(line[k.end():], bold_prefix=f"Art. {label(n)}{'.' if n >= 10 else ''} ")
        elif line.startswith("CL["):
            k = re.match(r"CL\[(\w+)\]\s*", line)
            p = para(f"CLÁUSULA {clauses[k.group(1)]}ª – {line[k.end():]}", bold=True)
            p.paragraph_format.space_before = Pt(10)
        elif line.startswith("%PAGEBREAK"):
            doc.add_page_break()
        elif line.startswith("P "):
            p = para(line[2:]); p.paragraph_format.first_line_indent = Cm(1.25)
        elif line.startswith("%SIGNS "):
            signatures([tuple(c.split(";")) if ";" in c else (None, c) for c in line[7:].split("|")])
        elif line.startswith("%SIGN"):
            date_line()
            signatures([("[NOME DO PRESIDENTE DA ASSEMBLEIA]", "Presidente da Assembleia"),
                        ("[NOME DO SECRETÁRIO DA ASSEMBLEIA]", "Secretário da Assembleia"),
                        ("[NOME DO ADVOGADO] – OAB/[UF] nº [NÚMERO]",
                         "Visto do advogado (Lei nº 8.906/1994, art. 1º, § 2º)")])
        elif line.startswith("%DATE"):
            date_line()
        elif line.startswith("%TABLE "):
            rows, cols = line[7:].split(" ", 1)
            cols = cols.split("|")
            t = doc.add_table(rows=int(rows) + 1, cols=len(cols))
            t.style = "Table Grid"
            for i, c in enumerate(cols):
                cell = t.rows[0].cells[i]
                cell.text = ""
                r = cell.paragraphs[0].add_run(c); r.bold = True; r.font.size = Pt(10)
            for i in range(1, int(rows) + 1):
                t.rows[i].cells[0].text = str(i)
                t.rows[i].height = Cm(0.9)
            widths = [1.0, 5.0, 3.0, 3.5, 3.5][:len(cols)]
            for row in t.rows:
                for cell, w in zip(row.cells, widths):
                    cell.width = Cm(w)
        elif re.match(r"(§|Parágrafo único)", line):
            m = re.match(r"(§ \d+º(?:-[A-Z])?|§ \d+\.|Parágrafo único\.)\s*", line)
            para(line[m.end():], bold_prefix=m.group(1) + " ", indent=1)
        else:  # incisos e listas
            para(line, indent=1.5 if re.match(r"[IVXL]+ –", line) else 1)

    out = base / out_name
    doc.save(out)
    print(out.name)

for src_name, out_name in DOCS:
    build(src_name, out_name)
print(f"{len(nums)} artigos no estatuto")
