import io
from datetime import date
from zipfile import ZIP_DEFLATED, ZipFile
from pathlib import Path

import streamlit as st
from lxml import etree
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

HEADER_IMAGE_PATH = Path(__file__).with_name("header_full.png")
FONT_NAME = "標楷體"


def _set_run_font(run):
  """同時指定中西文字型，避免 Word 只替換英文或只替換中文。"""
  run.font.name = FONT_NAME
  r_pr = run._element.get_or_add_rPr()
  r_fonts = r_pr.rFonts
  if r_fonts is None:
    r_fonts = OxmlElement("w:rFonts")
    r_pr.insert(0, r_fonts)
  for font_attr in ("ascii", "hAnsi", "eastAsia", "cs"):
    r_fonts.set(qn(f"w:{font_attr}"), FONT_NAME)


def _iter_table_paragraphs(table):
  for row in table.rows:
    for cell in row.cells:
      yield from cell.paragraphs
      for nested_table in cell.tables:
        yield from _iter_table_paragraphs(nested_table)


def _iter_all_paragraphs(doc):
  """涵蓋內文、表格、頁首與頁尾中的全部段落。"""
  yield from doc.paragraphs
  for table in doc.tables:
    yield from _iter_table_paragraphs(table)

  seen_parts = set()
  for section in doc.sections:
    for container in (
        section.header,
        section.first_page_header,
        section.even_page_header,
        section.footer,
        section.first_page_footer,
        section.even_page_footer,
    ):
      part_name = str(container.part.partname)
      if part_name in seen_parts:
        continue
      seen_parts.add(part_name)
      yield from container.paragraphs
      for table in container.tables:
        yield from _iter_table_paragraphs(table)


def _apply_font_everywhere(doc):
  """強制所有樣式與所有現有文字使用標楷體，保留原本字級及粗斜體。"""
  for style in doc.styles:
    if hasattr(style, "font"):
      style.font.name = FONT_NAME
      style_r_pr = style.element.get_or_add_rPr()
      style_r_fonts = style_r_pr.rFonts
      if style_r_fonts is None:
        style_r_fonts = OxmlElement("w:rFonts")
        style_r_pr.insert(0, style_r_fonts)
      for font_attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        style_r_fonts.set(qn(f"w:{font_attr}"), FONT_NAME)

  for paragraph in _iter_all_paragraphs(doc):
    for run in paragraph.runs:
      _set_run_font(run)


def _force_font_in_docx_package(docx_stream):
  """連註腳、文字方塊與條件式表格樣式也一併強制改成標楷體。"""
  source = io.BytesIO(docx_stream.getvalue())
  destination = io.BytesIO()
  word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  ns = {"w": word_ns}

  with ZipFile(source, "r") as input_zip:
    with ZipFile(destination, "w", ZIP_DEFLATED) as output_zip:
      for item in input_zip.infolist():
        data = input_zip.read(item.filename)
        if item.filename.startswith("word/") and item.filename.endswith(".xml"):
          root = etree.fromstring(data)
          for r_pr in root.xpath(".//w:rPr", namespaces=ns):
            r_fonts = r_pr.find(f"{{{word_ns}}}rFonts")
            if r_fonts is None:
              r_fonts = etree.Element(f"{{{word_ns}}}rFonts")
              r_pr.insert(0, r_fonts)
            for font_attr in ("ascii", "hAnsi", "eastAsia", "cs"):
              r_fonts.set(f"{{{word_ns}}}{font_attr}", FONT_NAME)
          data = etree.tostring(
              root, xml_declaration=True, encoding="UTF-8", standalone=True
          )
        output_zip.writestr(item, data)

  destination.seek(0)
  return destination


def _send_picture_to_back(run):
  """把行內圖片改為固定在頁面最下層的浮動背景圖。"""
  drawing = run._r.xpath("./w:drawing")[0]
  inline = drawing.xpath("./wp:inline")[0]
  inline.tag = qn("wp:anchor")
  for name, value in {
      "distT": "0",
      "distB": "0",
      "distL": "0",
      "distR": "0",
      "simplePos": "0",
      "relativeHeight": "0",
      "behindDoc": "1",
      "locked": "0",
      "layoutInCell": "1",
      "allowOverlap": "1",
  }.items():
    inline.set(name, value)

  extent = inline.find(qn("wp:extent"))
  simple_pos = OxmlElement("wp:simplePos")
  simple_pos.set("x", "0")
  simple_pos.set("y", "0")

  position_h = OxmlElement("wp:positionH")
  position_h.set("relativeFrom", "page")
  position_h.append(OxmlElement("wp:posOffset"))
  position_h[0].text = "0"

  position_v = OxmlElement("wp:positionV")
  position_v.set("relativeFrom", "page")
  position_v.append(OxmlElement("wp:posOffset"))
  position_v[0].text = "0"

  insert_at = list(inline).index(extent)
  inline.insert(insert_at, simple_pos)
  inline.insert(insert_at + 1, position_h)
  inline.insert(insert_at + 2, position_v)

  doc_pr = inline.find(qn("wp:docPr"))
  wrap_none = OxmlElement("wp:wrapNone")
  inline.insert(list(inline).index(doc_pr), wrap_none)


def apply_template_and_format(doc_stream, header_title, header_date):
  """套用完整橫幅背景，並疊加可編輯的速報標題與日期。"""
  doc = Document(doc_stream)
  if not HEADER_IMAGE_PATH.exists():
    raise FileNotFoundError(
        f"找不到固定頁首圖片：{HEADER_IMAGE_PATH.name}"
    )
  header_image = HEADER_IMAGE_PATH.read_bytes()

  # 1. 完整橫幅不拆分，固定在每頁頁首的最下層。
  doc.settings.odd_and_even_pages_header_footer = False
  for section in doc.sections:
    section.different_first_page_header_footer = False
    section.header.is_linked_to_previous = False
    section.header_distance = Inches(0)
    if section.top_margin < Inches(2.10):
      section.top_margin = Inches(2.10)

    header = section.header
    for child in list(header._element):
      header._element.remove(child)

    background_paragraph = header.add_paragraph()
    background_paragraph.paragraph_format.space_before = Pt(0)
    background_paragraph.paragraph_format.space_after = Pt(0)
    background_paragraph.paragraph_format.line_spacing = Pt(1)
    background_run = background_paragraph.add_run()
    background_run.add_picture(
        io.BytesIO(header_image), width=section.page_width
    )
    _send_picture_to_back(background_run)

    # 這兩段是真正的 Word 文字，下載後仍可直接修改。
    title_paragraph = header.add_paragraph()
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.paragraph_format.space_before = Pt(52)
    title_paragraph.paragraph_format.space_after = Pt(0)
    title_run = title_paragraph.add_run(header_title.strip())
    title_run.font.name = FONT_NAME
    title_run.font.size = Pt(16)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(0, 0, 0)

    date_paragraph = header.add_paragraph()
    date_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    date_paragraph.paragraph_format.right_indent = Inches(0.65)
    date_paragraph.paragraph_format.space_before = Pt(12)
    date_paragraph.paragraph_format.space_after = Pt(0)
    date_run = date_paragraph.add_run(header_date.strip())
    date_run.font.name = FONT_NAME
    date_run.font.size = Pt(12)
    date_run.font.color.rgb = RGBColor(0, 35, 166)

  # 2. 調整內文全域樣式（統一標楷體）
  normal_style = doc.styles["Normal"]
  normal_style.font.name = FONT_NAME
  normal_style.font.size = Pt(12)
  normal_style.font.color.rgb = RGBColor(51, 51, 51)
  normal_style.paragraph_format.line_spacing = 1.3
  normal_style.paragraph_format.space_after = Pt(6)

  # 3. 逐段巡邏並格式化標題（標楷體 + 加粗）
  for p in doc.paragraphs:
    if p.text.strip():
      if (
          p.text.startswith("一、")
          or p.text.startswith("二、")
          or p.text.startswith("三、")
          or len(p.text) < 20 and not p.text.endswith("。")
      ):
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        for run in p.runs:
          run.font.name = FONT_NAME
          run.font.size = Pt(14)
          run.font.bold = True
          run.font.color.rgb = RGBColor(0, 35, 102)
      else:
        for run in p.runs:
          run.font.name = FONT_NAME
          run.font.size = Pt(12)

  # 4. 內文圖片自動置中與縮放
  for p in doc.paragraphs:
    if "drawing" in p._p.xml:
      p.alignment = WD_ALIGN_PARAGRAPH.CENTER
      for run in p.runs:
        for inline in run._r.xpath(".//wp:inline"):
          inline.width = Inches(5.5)

  # 最後再做一次全面字型覆蓋，確保表格、頁首頁尾及所有樣式都一致。
  _apply_font_everywhere(doc)

  # 儲存到記憶體
  output_stream = io.BytesIO()
  doc.save(output_stream)
  output_stream.seek(0)
  return _force_font_in_docx_package(output_stream)


# --- Streamlit 網頁介面設計 ---
st.set_page_config(
    page_title="合庫標準範本自動排版系統", page_icon="📄", layout="centered"
)

st.title("📄 合庫標準報告範本自動套用工具")
st.write(
    "上傳原始 Word 檔案後，系統會把指定的完整合庫橫幅圖片直接放入頁首，"
    "不拆分任何元素並固定在背景最下層；中央標題及右下日期為可編輯文字。"
)
st.markdown("---")

# 檔案上傳區
today = date.today()
roc_date = f"日期：{today.year - 1911} 年 {today.month:02d} 月 {today.day:02d} 日"

uploaded_doc = st.file_uploader(
    "1. 請上傳您的原始 Word 檔案 (.docx)", type=["docx"]
)
header_title = st.text_input(
    "2. 速報標題（顯示在圖片中央，下載 Word 後仍可編輯）",
    value="私募信貸危機：起因、進展與風險",
)
header_date = st.text_input(
    "3. 日期（顯示在圖片右下方，下載 Word 後仍可編輯）",
    value=roc_date,
)

if uploaded_doc is not None:
  if st.button("🚀 開始自動套用速報範本與排版"):
    with st.spinner("正在處理中，請稍候..."):
      try:
        processed_doc = apply_template_and_format(
            uploaded_doc, header_title, header_date
        )

        st.success("🎉 完整範本與標楷體排版套用成功！")
        st.download_button(
            label="📥 下載合庫標準範本 Word 檔案",
            data=processed_doc,
            file_name="合庫標準範本_正式報告.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )
      except Exception as e:
        st.error(f"發生錯誤：{e}")