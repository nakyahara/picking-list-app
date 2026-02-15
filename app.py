# -*- coding: utf-8 -*-
"""
ピッキングリスト突合アプリ
TMP1（トータルピッキングリスト）に納品プランNoを追記するWebアプリ

入力: TMP1 PDF + マッピングCSV（GASが自動出力）
出力: 納品プランNo追記済みTMP1 PDF
"""

import streamlit as st
import pdfplumber
import csv
import io
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, white, black
import os
from datetime import datetime

# --- ページ設定 ---
st.set_page_config(
    page_title="ピッキングリスト突合ツール",
    page_icon="📋",
    layout="wide",
)

# --- 日本語フォント登録 ---
@st.cache_resource
def register_font():
    """日本語フォントを登録する"""
    font_name = "Japanese"
    local_fonts = [
        "C:/Windows/Fonts/msgothic.ttc",
        "C:/Windows/Fonts/meiryo.ttc",
        "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-JP-Regular.otf",
        "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
    ]
    for fp in local_fonts:
        if os.path.exists(fp):
            try:
                pdfmetrics.registerFont(TTFont(font_name, fp))
                return font_name
            except:
                pass

    return "Helvetica"


def read_mapping_csv(csv_file):
    """マッピングCSVから商品ID→納品プランNoの辞書を作成"""
    plan_map = {}
    csv_file.seek(0)
    raw = csv_file.read()

    # エンコーディング自動判定（UTF-8 → CP932）
    for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
        try:
            text = raw.decode(enc)
            break
        except (UnicodeDecodeError, AttributeError):
            continue
    else:
        text = raw.decode('utf-8', errors='replace')

    reader = csv.reader(io.StringIO(text))
    header = next(reader, None)
    if not header:
        return plan_map

    # ヘッダーから商品IDと納品プランNoの列を検出
    id_col = None
    plan_col = None
    for i, h in enumerate(header):
        h_clean = h.strip()
        if h_clean in ('商品ID', '商品コード'):
            id_col = i
        elif h_clean in ('納品プランNo', 'プランNo', 'ラベル'):
            plan_col = i

    # ヘッダーが見つからない場合、最初の2列を使用
    if id_col is None or plan_col is None:
        id_col = 0
        plan_col = 1

    for row in reader:
        if len(row) > max(id_col, plan_col):
            pid = row[id_col].strip()
            plan_no = row[plan_col].strip()
            if pid and plan_no:
                # 複数プランが改行区切りの場合は「/」区切りに変換して全て保持
                plans = [p.strip() for p in plan_no.replace('\r', '').split('\n') if p.strip()]
                if plans:
                    plan_map[pid] = ' / '.join(plans)

    return plan_map


def extract_tmp1_page_data(pdf_file):
    """TMP1の各ページから商品IDとセル位置情報を抽出"""
    page_data = []

    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables_found = page.find_tables()
            if not tables_found:
                page_data.append([])
                continue

            data_table = tables_found[-1]
            rows = data_table.rows

            extracted = page.extract_tables()
            ext_table = extracted[-1] if extracted else []

            items = []
            ri = 3  # ヘッダー3行スキップ
            while ri < len(rows):
                row = rows[ri]
                cells = row.cells
                if cells and cells[0]:
                    y_top = cells[0][1]
                    y_bottom = cells[0][3]

                    product_id = ''
                    if ri < len(ext_table) and ext_table[ri][1]:
                        product_id = ext_table[ri][1].strip()

                    items.append({
                        'product_id': product_id,
                        'y_top': y_top,
                        'y_bottom': y_bottom,
                    })
                    ri += 3
                else:
                    ri += 1

            page_data.append(items)

    return page_data


def create_merged_pdf(tmp1_file, plan_map, page_data, font_name):
    """元のTMP1 PDFに納品プランNo列をオーバーレイして新しいPDFを作成"""
    EXTRA_WIDTH = 75

    reader = PdfReader(tmp1_file)
    writer = PdfWriter()

    matched_count = 0
    unmatched_ids = set()

    for page_num, orig_page in enumerate(reader.pages):
        mb = orig_page.mediabox
        orig_width = float(mb.width)
        orig_height = float(mb.height)
        new_width = orig_width + EXTRA_WIDTH

        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(new_width, orig_height))

        items = page_data[page_num] if page_num < len(page_data) else []

        table_right = 571.8
        col_x = table_right
        col_width = EXTRA_WIDTH

        def to_pdf_y(plumber_y):
            return orig_height - plumber_y

        if items:
            # ヘッダー
            header_top = 117.9
            header_bottom = 160.4
            hy_top_pdf = to_pdf_y(header_top)
            hy_bottom_pdf = to_pdf_y(header_bottom)

            c.setFillColor(HexColor('#808080'))
            c.setStrokeColor(HexColor('#808080'))
            c.rect(col_x, hy_bottom_pdf, col_width, hy_top_pdf - hy_bottom_pdf, fill=1, stroke=1)

            c.setFillColor(white)
            c.setFont(font_name, 7)
            text_y = hy_bottom_pdf + (hy_top_pdf - hy_bottom_pdf) / 2 + 8
            c.drawCentredString(col_x + col_width / 2, text_y, "納品プラン")
            c.drawCentredString(col_x + col_width / 2, text_y - 12, "No")

            # データ行
            c.setStrokeColor(HexColor('#808080'))
            c.setLineWidth(0.5)

            for item in items:
                y_top_pdf = to_pdf_y(item['y_top'])
                y_bottom_pdf = to_pdf_y(item['y_bottom'])
                cell_height = y_top_pdf - y_bottom_pdf

                c.setFillColor(HexColor('#FFFFFF'))
                c.rect(col_x, y_bottom_pdf, col_width, cell_height, fill=1, stroke=1)

                pid = item['product_id']
                plan_no = plan_map.get(pid, '')
                if plan_no:
                    matched_count += 1
                else:
                    plan_no = '(該当なし)'
                    if pid:
                        unmatched_ids.add(pid)

                c.setFillColor(black)
                # 複数プランの場合はフォント縮小
                font_size = 5.5 if ' / ' in plan_no else 7
                c.setFont(font_name, font_size)
                text_y = y_bottom_pdf + cell_height / 2 - 3
                c.drawCentredString(col_x + col_width / 2, text_y, plan_no)

        c.save()
        packet.seek(0)

        overlay_reader = PdfReader(packet)
        overlay_page = overlay_reader.pages[0]

        orig_page.mediabox.upper_right = (new_width, orig_height)
        if "/CropBox" in orig_page:
            orig_page.cropbox.upper_right = (new_width, orig_height)
        orig_page.merge_page(overlay_page)
        writer.add_page(orig_page)

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)

    return output, matched_count, unmatched_ids


# =============================================
# メイン画面
# =============================================

st.title("📋 ピッキングリスト突合ツール")
st.markdown("トータルピッキングリスト（TMP1）に**納品プランNo**を自動追記します。")

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("① トータルピッキングリスト")
    st.caption("ロジザードから出力したPDF（TMP1）")
    tmp1_file = st.file_uploader(
        "TMP1をアップロード",
        type=["pdf"],
        key="tmp1",
        label_visibility="collapsed",
    )

with col2:
    st.subheader("② マッピングCSV")
    st.caption("GASが自動出力する商品ID→納品プランNoのCSV")
    mapping_file = st.file_uploader(
        "マッピングCSVをアップロード",
        type=["csv"],
        key="mapping",
        label_visibility="collapsed",
    )

st.divider()

# 実行ボタン
if tmp1_file and mapping_file:
    if st.button("🔄 突合実行", type="primary", use_container_width=True):
        with st.spinner("処理中..."):
            font_name = register_font()

            # Step 1: CSVからマッピング読み取り
            progress = st.progress(0, text="マッピングCSVを読み込み中...")
            plan_map = read_mapping_csv(mapping_file)
            progress.progress(30, text=f"納品プランNoマッピング: {len(plan_map)}件")

            # Step 2: TMP1解析
            progress.progress(40, text="トータルピッキングリストを解析中...")
            tmp1_file.seek(0)
            page_data = extract_tmp1_page_data(tmp1_file)
            total_items = sum(len(items) for items in page_data)
            progress.progress(60, text=f"TMP1データ: {total_items}件")

            # Step 3: PDF生成
            progress.progress(70, text="PDFを生成中...")
            tmp1_file.seek(0)
            result_pdf, matched, unmatched_ids = create_merged_pdf(
                tmp1_file, plan_map, page_data, font_name
            )
            progress.progress(100, text="完了！")

            # 結果表示
            st.success(f"突合完了！ {matched}/{total_items}件 マッチしました。")

            if unmatched_ids:
                with st.expander(f"⚠️ 該当なし: {len(unmatched_ids)}件（クリックで詳細）"):
                    st.markdown("以下の商品IDはマッピングCSVに見つかりませんでした：")
                    for uid in sorted(unmatched_ids):
                        st.code(uid)

            # ダウンロードボタン
            today = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"TMP1_納品プランNo追記済_{today}.pdf"

            st.download_button(
                label="📥 追記済PDFをダウンロード",
                data=result_pdf,
                file_name=filename,
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )
else:
    st.info("👆 2つのファイルをアップロードしてください。")


# フッター
st.divider()
st.caption("ピッキングリスト突合ツール v2.0 | CSVベースで100%正確なマッピング")
