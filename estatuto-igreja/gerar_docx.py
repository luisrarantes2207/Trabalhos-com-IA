"""Gera o .docx do estatuto a partir de estatuto.txt (numeração automática de artigos)."""
import re, sys
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

base = Path(__file__).parent
src = (base / "estatuto.txt").read_text(encoding="utf-8").splitlines()

def label(n):
    return f"{n}º" if n < 10 else f"{n}"

# 1ª passada: numeração
nums, n = {}, 0
for line in src:
    m = re.match(r"ART\[(\w+)\]", line)
    if m:
        n += 1
        nums[m.group(1)] = n

def refs(text):
    def sub(m):
        k = m.group(1)
        if k not in nums:
            sys.exit(f"Referência inexistente: {k}")
        return label(nums[k])
    return re.sub(r"\[\[(\w+)\]\]", sub, text)

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

for line in src:
    line = refs(line.rstrip())
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
    elif line.startswith("ART["):
        k = re.match(r"ART\[(\w+)\]\s*", line)
        n = nums[k.group(1)]
        para(line[k.end():], bold_prefix=f"Art. {label(n)}{'.' if n >= 10 else ''} ")
    elif line.startswith("%SIGN"):
        para("[CIDADE]/[UF], [DIA] de [MÊS] de [ANO].", align=WD_ALIGN_PARAGRAPH.RIGHT).paragraph_format.space_before = Pt(18)
        for nome, cargo in [("[NOME DO PRESIDENTE DA ASSEMBLEIA]", "Presidente da Assembleia"),
                            ("[NOME DO SECRETÁRIO DA ASSEMBLEIA]", "Secretário da Assembleia"),
                            ("[NOME DO ADVOGADO] – OAB/[UF] nº [NÚMERO]",
                             "Visto do advogado (Lei nº 8.906/1994, art. 1º, § 2º)")]:
            p = para("__________________________________________", align=WD_ALIGN_PARAGRAPH.CENTER)
            p.paragraph_format.space_before = Pt(30); p.paragraph_format.space_after = Pt(0)
            p = para(nome, align=WD_ALIGN_PARAGRAPH.CENTER); p.paragraph_format.space_after = Pt(0)
            para(cargo, align=WD_ALIGN_PARAGRAPH.CENTER)
    elif re.match(r"(§|Parágrafo único)", line):
        m = re.match(r"(§ \d+º(?:-[A-Z])?|§ \d+\.|Parágrafo único\.)\s*", line)
        para(line[m.end():], bold_prefix=m.group(1) + " ", indent=1)
    else:  # incisos
        para(line, indent=1.5 if re.match(r"[IVXL]+ –", line) else 1)

out = base / "Estatuto_Social_Igreja_ISAC_v2.docx"
doc.save(out)
print(f"{out} — {len(nums)} artigos")
