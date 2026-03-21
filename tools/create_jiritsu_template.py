"""
自立支援医療診断書（精神通院医療用）のWordテンプレート (.docx) を生成する
一度だけ実行して sample/自立支援/jiritsu_shien_template.docx を作成する
"""
import os
from docx import Document
from docx.shared import Pt, Mm, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

OUTPUT = os.path.join(os.path.dirname(__file__), "..", "sample", "自立支援", "jiritsu_shien_template.docx")

def set_cell_shading(cell, color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}></w:tcBorders>')
    for edge, val in kwargs.items():
        element = parse_xml(
            f'<w:{edge} {nsdecls("w")} w:val="{val.get("val","single")}" '
            f'w:sz="{val.get("sz","4")}" w:space="0" w:color="{val.get("color","000000")}"/>'
        )
        tcBorders.append(element)
    tcPr.append(tcBorders)

def make_label_cell(cell, text, font_size=8, bold=True):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.font.size = Pt(font_size)
    run.font.name = "ＭＳ 明朝"
    run.bold = bold
    set_cell_shading(cell, "F0F0F0")
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def make_value_cell(cell, text="", font_size=9):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.font.size = Pt(font_size)
    run.font.name = "ＭＳ 明朝"
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def add_section_header(table, text, cols=1):
    row = table.add_row()
    cell = row.cells[0]
    if cols > 1:
        cell = row.cells[0].merge(row.cells[cols - 1])
    make_label_cell(cell, text, font_size=9, bold=True)

def add_value_row(table, text="", cols=1, min_height_cm=None):
    row = table.add_row()
    cell = row.cells[0]
    if cols > 1:
        cell = row.cells[0].merge(row.cells[cols - 1])
    make_value_cell(cell, text, font_size=9)
    if min_height_cm:
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        trHeight = parse_xml(
            f'<w:trHeight {nsdecls("w")} w:val="{int(min_height_cm * 567)}" w:hRule="atLeast"/>'
        )
        trPr.append(trHeight)
    return cell

def set_row_height(row, cm):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = parse_xml(
        f'<w:trHeight {nsdecls("w")} w:val="{int(cm * 567)}" w:hRule="atLeast"/>'
    )
    trPr.append(trHeight)

def main():
    doc = Document()

    # ページ設定 (A4)
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    style = doc.styles["Normal"]
    style.font.name = "ＭＳ 明朝"
    style.font.size = Pt(9)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.line_spacing = Pt(14)

    # ===== タイトル =====
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("自立支援医療診断書（精神通院医療用）")
    run.font.size = Pt(14)
    run.font.name = "ＭＳ 明朝"
    run.bold = True

    doc.add_paragraph()  # spacer

    # ===== Table 0: 患者基本情報 =====
    t0 = doc.add_table(rows=5, cols=4)
    t0.style = "Table Grid"
    t0.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Row 0: フリガナ
    make_label_cell(t0.cell(0, 0), "フリガナ")
    t0.cell(0, 1).merge(t0.cell(0, 3))
    make_value_cell(t0.cell(0, 1), "")

    # Row 1: 氏名
    make_label_cell(t0.cell(1, 0), "氏　名")
    t0.cell(1, 1).merge(t0.cell(1, 3))
    make_value_cell(t0.cell(1, 1), "")

    # Row 2: 生年月日・年齢 | 性別
    make_label_cell(t0.cell(2, 0), "生年月日")
    make_value_cell(t0.cell(2, 1), "　　年　　月　　日生（　　歳）")
    make_label_cell(t0.cell(2, 2), "性別")
    make_value_cell(t0.cell(2, 3), "男　・　女")

    # Row 3: 住所
    make_label_cell(t0.cell(3, 0), "住　所")
    t0.cell(3, 1).merge(t0.cell(3, 3))
    make_value_cell(t0.cell(3, 1), "〒")
    set_row_height(t0.rows[3], 1.2)

    # Row 4: 連絡先電話番号
    make_label_cell(t0.cell(4, 0), "電話番号")
    t0.cell(4, 1).merge(t0.cell(4, 3))
    make_value_cell(t0.cell(4, 1), "")

    doc.add_paragraph()

    # ===== Table 1: 1 病名 =====
    t1 = doc.add_table(rows=1, cols=4)
    t1.style = "Table Grid"
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header
    hdr = t1.cell(0, 0).merge(t1.cell(0, 3))
    make_label_cell(hdr, "１　病名　（ICD-10コードはF00〜F99、G40のいずれかを記載）")

    # (1) 主たる精神障害
    r = t1.add_row()
    make_label_cell(r.cells[0], "(1) 主たる精神障害", font_size=8)
    make_value_cell(r.cells[1], "")
    make_label_cell(r.cells[2], "ICD-10", font_size=8)
    make_value_cell(r.cells[3], "")

    # (2) 従たる精神障害
    r = t1.add_row()
    make_label_cell(r.cells[0], "(2) 従たる精神障害", font_size=8)
    make_value_cell(r.cells[1], "")
    make_label_cell(r.cells[2], "ICD-10", font_size=8)
    make_value_cell(r.cells[3], "")

    # (3) 身体合併症
    r = t1.add_row()
    make_label_cell(r.cells[0], "(3) 身体合併症", font_size=8)
    c = r.cells[1].merge(r.cells[3])
    make_value_cell(c, "")

    doc.add_paragraph()

    # ===== Table 2: 2 発病から現在までの病歴 =====
    t2 = doc.add_table(rows=1, cols=4)
    t2.style = "Table Grid"
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr = t2.cell(0, 0).merge(t2.cell(0, 3))
    make_label_cell(hdr, "２　発病から現在までの病歴")

    # 推定発病時期・初診日
    r = t2.add_row()
    make_label_cell(r.cells[0], "推定発病時期", font_size=8)
    make_value_cell(r.cells[1], "　　年　　月頃")
    make_label_cell(r.cells[2], "初診日", font_size=8)
    make_value_cell(r.cells[3], "　　年　　月　　日")

    # 病歴本文
    r = t2.add_row()
    c = r.cells[0].merge(r.cells[3])
    make_value_cell(c, "")
    set_row_height(r, 6.0)

    doc.add_paragraph()

    # ===== Table 3: 3 現在の病状、状態像等 =====
    t3 = doc.add_table(rows=1, cols=1)
    t3.style = "Table Grid"
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER

    make_label_cell(t3.cell(0, 0), "３　現在の病状、状態像等（該当する項目を○で囲んでください）")

    symptom_items = [
        "(1) そう状態\n"
        "　　１思考奔逸・誇大妄想　２自殺願望・自殺企図　３カタレプシー　４その他（　　　　　　　　　　）",
        "(2) 抑うつ状態\n"
        "　　１思考・運動抑制　２易刺激性・興奮　３憂うつ気分　４その他（　　　　　　　　　　）",
        "(3) 幻覚妄想状態\n"
        "　　１幻覚　２妄想　３その他（　　　　　　　　　　）",
        "(4) 精神運動興奮及び昏迷・拒絶\n"
        "　　１興奮　２昏迷　３拒絶　４その他（　　　　　　　　　　）",
        "(5) 残遺状態又は精神欠損状態\n"
        "　　１自閉　２感情平板化　３意欲の減退　４その他（　　　　　　　　　　）",
        "(6) 情緒及び行動の障害\n"
        "　　１爆発性　２易怒性　３気分変動　４暴力・衝動行為　５常同行為　６多動\n"
        "　　７食行動の異常　８性行動の異常　９チック・汚言　10その他（　　　　　　　　　　）",
        "(7) 不安及び不穏状態\n"
        "　　１強度の不安・恐怖感　２精神運動不穏　３心身衰弱　４強迫体験　５心気症状\n"
        "　　６心的外傷に関連する症状　７解離・転換症状　８その他（　　　　　　　　　　）",
        "(8) てんかん発作等（意識障害を含む）\n"
        "　　１てんかん発作　発作型（イ・ロ・ハ・ニ）　頻度（　　回/年）　最終発作（　　年　　月　　日）\n"
        "　　２意識障害　３その他（　　　　　　　　　　）",
        "(9) 精神作用物質の乱用・依存\n"
        "　　１アルコール　２覚醒剤　３有機溶剤　４その他（　　　　　　　　　　）\n"
        "　　ア乱用　イ依存　ウ残遺性・遅発性精神病性障害　エその他（　　　　　　　　　　）",
        "(10) 知能、記憶、学習の障害\n"
        "　　１知的障害（ア軽度　イ中等度　ウ重度）　２認知症　３その他の記憶障害（　　　　　　）\n"
        "　　４学習の困難（ア読み　イ書き　ウ算数　エその他）　５遂行機能障害　６注意障害　７その他（　　　）",
        "(11) 広汎性発達障害関連症状\n"
        "　　１相互的な社会関係の質的障害　２コミュニケーションのパターンにおける質的障害\n"
        "　　３限定した常同的で反復的な関心と活動　４その他（　　　　　　　　　　）",
        "(12) その他（　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　）",
    ]

    for item in symptom_items:
        r = t3.add_row()
        make_value_cell(r.cells[0], item, font_size=8)

    doc.add_paragraph()

    # ===== Table 4: 4 具体的程度 =====
    t4 = doc.add_table(rows=1, cols=1)
    t4.style = "Table Grid"
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER

    make_label_cell(t4.cell(0, 0),
                    "４　３の病状・状態像の具体的程度、症状、処方薬等\n"
                    "（現在の病状が日常生活への影響を含め具体的にお書きください）")
    r = t4.add_row()
    make_value_cell(r.cells[0], "")
    set_row_height(r, 5.0)

    doc.add_paragraph()

    # ===== Table 5: 5 現在の治療内容 =====
    t5 = doc.add_table(rows=1, cols=1)
    t5.style = "Table Grid"
    t5.alignment = WD_TABLE_ALIGNMENT.CENTER

    make_label_cell(t5.cell(0, 0), "５　現在の治療内容")

    # (1) 投薬内容
    r = t5.add_row()
    make_value_cell(r.cells[0],
                    "（１）投薬内容（※診断名に対する薬剤名（商品名可）をご記載ください。）\n",
                    font_size=8)
    set_row_height(r, 4.0)

    # (2) 精神療法等
    r = t5.add_row()
    make_value_cell(r.cells[0],
                    "（２）精神療法等\n",
                    font_size=8)
    set_row_height(r, 3.5)

    # (3) 訪問看護指示の有無
    r = t5.add_row()
    make_value_cell(r.cells[0],
                    "（３）訪問看護指示の有無（　有　・　無　）",
                    font_size=8)

    doc.add_paragraph()

    # ===== Table 6: 6 今後の治療方針 =====
    t6 = doc.add_table(rows=1, cols=1)
    t6.style = "Table Grid"
    t6.alignment = WD_TABLE_ALIGNMENT.CENTER

    make_label_cell(t6.cell(0, 0), "６　今後の治療方針")
    r = t6.add_row()
    make_value_cell(r.cells[0], "")
    set_row_height(r, 4.0)

    doc.add_paragraph()

    # ===== Table 7: 7 障害福祉サービス =====
    t7 = doc.add_table(rows=1, cols=1)
    t7.style = "Table Grid"
    t7.alignment = WD_TABLE_ALIGNMENT.CENTER

    make_label_cell(t7.cell(0, 0), "７　現在の障害福祉サービス等の利用状況（該当する項目を○で囲んでください）")

    r = t7.add_row()
    make_value_cell(r.cells[0],
                    "（１）自立訓練（生活訓練）\n"
                    "（２）共同生活援助（グループホーム）\n"
                    "（３）居宅介護（ホームヘルプ）\n"
                    "（４）その他の障害福祉サービス（　　　　　　　　　　）\n"
                    "（５）訪問指導等\n"
                    "（６）なし",
                    font_size=8)

    doc.add_paragraph()

    # ===== Table 8: 8 備考 =====
    t8 = doc.add_table(rows=1, cols=1)
    t8.style = "Table Grid"
    t8.alignment = WD_TABLE_ALIGNMENT.CENTER

    make_label_cell(t8.cell(0, 0), "８　備考")
    r = t8.add_row()
    make_value_cell(r.cells[0], "")
    set_row_height(r, 2.5)

    doc.add_paragraph()

    # ===== Table 9: 署名欄 =====
    t9 = doc.add_table(rows=8, cols=2)
    t9.style = "Table Grid"
    t9.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Row 0: 日付
    make_label_cell(t9.cell(0, 0), "診断日")
    make_value_cell(t9.cell(0, 1), "　　年　　月　　日")

    # Row 1: 医療機関コード
    make_label_cell(t9.cell(1, 0), "医療機関コード")
    make_value_cell(t9.cell(1, 1), "")

    # Row 2: 所在地
    make_label_cell(t9.cell(2, 0), "医療機関所在地")
    make_value_cell(t9.cell(2, 1), "")

    # Row 3: 名称
    make_label_cell(t9.cell(3, 0), "医療機関名称")
    make_value_cell(t9.cell(3, 1), "")

    # Row 4: 電話番号
    make_label_cell(t9.cell(4, 0), "電話番号")
    make_value_cell(t9.cell(4, 1), "")

    # Row 5: 医師氏名
    make_label_cell(t9.cell(5, 0), "医師氏名")
    make_value_cell(t9.cell(5, 1), "")

    # Row 6: 指定医番号
    make_label_cell(t9.cell(6, 0), "精神保健指定医の証の番号")
    make_value_cell(t9.cell(6, 1), "")

    # Row 7: 従事年数
    make_label_cell(t9.cell(7, 0), "精神科従事年数")
    make_value_cell(t9.cell(7, 1), "　　　年")

    # ===== 重度かつ継続 判定欄 =====
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("※ 重度かつ継続の判定")
    run.font.size = Pt(8)
    run.font.name = "ＭＳ 明朝"
    run.bold = True

    # Table 10: 重度かつ継続
    t10 = doc.add_table(rows=2, cols=2)
    t10.style = "Table Grid"
    t10.alignment = WD_TABLE_ALIGNMENT.CENTER

    make_label_cell(t10.cell(0, 0), "該当")
    make_value_cell(t10.cell(0, 1), "該当　・　非該当")
    make_label_cell(t10.cell(1, 0), "理由")
    make_value_cell(t10.cell(1, 1), "")
    set_row_height(t10.rows[1], 1.5)

    # 注意書き
    doc.add_paragraph()
    note = doc.add_paragraph()
    run = note.add_run("※ 患者氏名等は空欄のまま（事務が後記入）")
    run.font.size = Pt(7)
    run.font.name = "ＭＳ 明朝"
    run.italic = True

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    doc.save(OUTPUT)
    print(f"テンプレート生成完了: {OUTPUT}")


if __name__ == "__main__":
    main()
