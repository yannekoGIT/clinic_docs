"""
自立支援医療診断書（精神通院医療用）Word版テンプレート記入スクリプト
公式書式 2jiritsu.docx の空欄にJSONデータを流し込む（レイアウト変更なし）

Usage: python jiritsu_shien_word.py <input.json> <output.docx>

テンプレート構造 (2jiritsu.docx):
  Table 0 (6 rows x 4 cols, 結合セル多数):
    R0: 氏名/フリガナ/生年月日
    R1: 住所
    R2: 1 病名 ((1)主, (2)従, (3)合併症 + ICDコード)
    R3: 2 病歴ヘッダー (推定発病時期/初診日)
    R4: 2 病歴本文 (空欄)
    R5: 3 現在の病状 (症状チェック P0-P33)
  Table 1 (8 rows x 3 cols, 結合セル多数):
    R0: 4 具体的程度 (P0-P1ヘッダー, P2空欄)
    R1: 5 治療内容ヘッダー + (1)投薬ラベル
    R2: 投薬値(P0-P2), (2)精神療法(P3-P4ラベル,P5-P7値), (3)訪問看護(P8)
    R3: 6 治療方針ヘッダー
    R4: 6 治療方針値 (空欄)
    R5: 7 福祉サービス(C0) | 8 備考ラベル(C2)
    R6: 7 福祉サービス(C0続) | 8 備考値(C2)
    R7: 日付/医療機関情報/医師
  Table 2: 行政記入欄 (変更不要)
"""
import sys, json, os, shutil
from docx import Document
from docx.shared import Pt
from lxml import etree
from docx.oxml.ns import nsdecls  # noqa: F401 (used in _fill_designated_info_xml)
from docx.oxml import parse_xml  # noqa: F401

TEMPLATE = os.path.join(os.path.dirname(__file__), "..", "..", "sample", "自立支援", "2jiritsu.docx")


def main():
    if len(sys.argv) < 3:
        print("Usage: python jiritsu_shien_word.py <input.json> <output.docx>", file=sys.stderr)
        sys.exit(1)
    data = json.loads(open(sys.argv[1], "r", encoding="utf-8").read())
    fill_template(data, sys.argv[2])
    print(f"生成完了: {sys.argv[2]}")


def fill_template(data, output_path):
    abs_output = os.path.abspath(output_path)
    shutil.copy2(TEMPLATE, abs_output)
    doc = Document(abs_output)

    t0 = doc.tables[0]
    t1 = doc.tables[1]  # python-docx: top-level tables only (Table 2 in XML = index 1)

    ins = _merge_clinic_defaults(data.get("institution", {}))
    cc = data.get("current_condition", {})
    treat = data.get("treatment", {})
    wf = data.get("welfare_services", {})
    pt = data.get("patient", {})

    # ======== Table 0 ========

    # --- R0: 氏名・フリガナ・生年月日 ---
    furigana = pt.get("furigana") or pt.get("name_kana") or pt.get("kana") or ""
    if furigana:
        _append_to_para(t0.cell(0, 1).paragraphs[0], "　" + furigana)
    if pt.get("name"):
        _write_para(t0.cell(0, 1).paragraphs[1], pt["name"])

    birthdate = pt.get("birthdate", "")
    age = pt.get("age")
    if age in (None, "") and birthdate:
        age = _calc_age(birthdate, data.get("date", ""))
    if birthdate:
        by, bm, bd = _parse_ymd(birthdate)
        era = _era_name(by)
        ey = _era_year(by)
        birth_text = f"{era}　{ey}年　{bm}月　{bd}日生（{age if age not in (None, '') else '　'}歳）"
        _write_para(t0.cell(0, 3).paragraphs[1], birth_text)

    # --- R1: 住所 ---
    address = _format_postal_address(
        pt.get("postal_code", "") or pt.get("zip_code", ""),
        pt.get("address", "")
    )
    if address:
        _write_para(t0.cell(1, 1).paragraphs[0], address)

    # --- R2: 1 病名 ---
    # Cell(2,2) = merged content area
    # P1 runs: R0(１) R1(主たる精神障害) R2(空) R3(空白19文字=値欄) R4(空) R5(ICD) R6(コード)
    # P3 runs: R0(２) R1(従たる精神障害) R2(空白=値欄) R3(空) R4(ICD) R5(コード)
    # P5: (3) 身体合併症 R4(空白=値欄)
    diag_cell = t0.cell(2, 2)
    if data.get("diagnosis_main"):
        _fill_diagnosis_run(diag_cell.paragraphs[1], 3, data["diagnosis_main"])
    if data.get("diagnosis_sub"):
        _fill_diagnosis_run(diag_cell.paragraphs[3], 2, data["diagnosis_sub"])
    if data.get("comorbidity"):
        _fill_diagnosis_run(diag_cell.paragraphs[5], 4, data["comorbidity"])

    # ICDコード → ネストテーブル (Table 0 R2 tc[1] 内の sub-table: 2行×6列)
    # Row 0 = 主たる精神障害のICD, Row 1 = 従たる精神障害のICD
    if data.get("icd_code_main") or data.get("icd_code_sub"):
        icd_nested = _find_nested_table(t0, 2, 1)  # Row 2, Cell 1 内のネストテーブル
        if icd_nested is not None:
            if data.get("icd_code_main"):
                _fill_icd_cells_xml(icd_nested, 0, data["icd_code_main"])
            if data.get("icd_code_sub"):
                _fill_icd_cells_xml(icd_nested, 1, data["icd_code_sub"])

    # --- R3: 2 病歴ヘッダー (推定発病時期/初診日) ---
    onset_y, onset_m = _parse_ym(data.get("onset_date", ""))
    fv_y, fv_m, fv_d = _parse_ymd(data.get("first_visit_date", ""))
    header_text = (
        f"（推定発病時期　{onset_y or '　　'}年{onset_m or '　'}月頃、"
        f"　初診日　{fv_y or '　　'}年{fv_m or '　'}月"
        f"{'　' + fv_d + '日' if fv_d else ''}）"
    )
    _write_para(t0.cell(3, 2).paragraphs[0], header_text)

    # --- R4: 2 病歴本文 ---
    course = data.get("clinical_course", "")
    if course:
        _write_para_compact(t0.cell(4, 2).paragraphs[0], course, max_chars=200)

    # --- R5: 3 症状チェックリスト ---
    _fill_symptoms(t0.cell(5, 0), cc)

    # ======== Table 1 ========

    # --- R0: 4 具体的程度 ---
    detail = data.get("condition_detail", "")
    if detail:
        _write_para_compact(t1.cell(0, 0).paragraphs[2], detail, max_chars=100)

    # --- R2: 5 治療内容 ---
    meds = treat.get("medications", [])
    med_cell = t1.cell(2, 0)

    # (1) 投薬内容 — 全薬剤を1段落に横並びで記載（高さ節約）
    if meds:
        med_parts = [f"{m['name']} {m.get('dosage','')} {m.get('frequency','')}" for m in meds]
        med_text = " / ".join(med_parts)
        _write_para(med_cell.paragraphs[0], med_text, font_size_override=Pt(8))

    # P5-P7: 精神療法値
    psych = treat.get("psychotherapy", "")
    if psych:
        _write_para_compact(med_cell.paragraphs[5], psych, max_chars=60)

    # P8: (3) 訪問看護指示 — テキストは中立のまま、括弧内の有/無に図形の○
    designated = treat.get("designated_doctor", "")
    if designated in ("有", "無"):
        nursing_text = "（３）訪問看護指示の有無（  有　・　無  ）"
        _write_para(med_cell.paragraphs[8], nursing_text)
        # "有無"の有/無ではなく括弧内の有/無をrfindで探す
        _circle_item_in_para(med_cell.paragraphs[8], designated,
                             include_number=False, use_last=True)

    # --- R4: 6 治療方針 ---
    plan = data.get("treatment_plan", "")
    if plan:
        _write_para_compact(t1.cell(4, 0).paragraphs[0], plan, max_chars=80)

    # --- R5-R6: 7 福祉サービス ---
    _fill_welfare(t1.cell(5, 0), wf)

    # --- R6 C2: 8 備考 ---
    remarks = data.get("remarks", "")
    if remarks:
        _write_para(t1.cell(6, 2).paragraphs[0], remarks)

    # --- R7: 日付・医療機関情報 ---
    # ※段落上書きではなく、ラベルを検索して値を追記する（注記や指定医情報を消さない）
    sig_cell = t1.cell(7, 0)
    d_y, d_m, d_d = _parse_ymd(data.get("date", ""))
    if d_y:
        _fill_label_value(sig_cell, "年", f"　　{d_y}年　{d_m or '　'}月　{d_d or '　'}日", exact_para_idx=1)

    # 医療機関コード → ネストテーブル (R7 tc[0] 内の sub-table: 1行×7列)
    code = ins.get("code", "")
    if code:
        code_nested = _find_nested_table(t1, 7, 0)
        if code_nested is not None:
            _fill_icd_cells_xml(code_nested, 0, code)

    # 医療機関所在地: ラベル "医療機関所在地" の後に住所を追記（〒除去）
    if ins.get("address"):
        addr = ins["address"].replace("〒", "").strip()
        _fill_label_value_small(sig_cell, "医療機関所在地", "　" + addr, Pt(9))

    # 名称: ラベル "称" を含むrunの後に名称を追記
    if ins.get("name"):
        _fill_label_value(sig_cell, "称", "　" + ins["name"])

    # 電話番号
    if ins.get("phone"):
        _fill_label_value(sig_cell, "電話番号", "　" + ins["phone"])

    # 医師氏名
    if ins.get("doctor"):
        _fill_label_value(sig_cell, "医師氏名", "　" + ins["doctor"])

    # 精神保健指定医の証の番号 / 精神医療従事年数
    # python-docxのAPIではアクセスできないため、XMLレベルで書き込む
    designated_num = ins.get("designated_number", "")
    years = ins.get("doctor_years", "")
    if designated_num or years:
        _fill_designated_info_xml(doc, designated_num, str(years) if years else "")

    doc.save(abs_output)


# ============================================================
# 図形の○（楕円シェイプ）ヘルパー — テキスト非変更で選択肢を囲む
# ============================================================

_shape_id_counter = 1000


def _next_shape_id():
    global _shape_id_counter
    _shape_id_counter += 1
    return _shape_id_counter


def _display_width(text):
    """全角=2, 半角=1 で表示幅を計算"""
    w = 0
    for ch in text:
        cp = ord(ch)
        if cp <= 0x7E or (0xFF61 <= cp <= 0xFF9F):
            w += 1
        else:
            w += 2
    return w


def _get_para_font_size_pt(para, default=10):
    """段落のフォントサイズ(pt)を取得"""
    for run in para.runs:
        if run.font.size:
            return run.font.size.pt
    return default


def _add_oval_to_para(para, left_pt, top_pt, width_pt, height_pt):
    """段落にフローティング楕円シェイプ（赤い○）を追加。
    Word上で図形を移動・削除するだけで修正できる。"""
    sid = _next_shape_id()
    to_emu = lambda pt: int(pt * 12700)

    xml = (
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<w:rPr><w:noProof/></w:rPr>'
        '<w:drawing>'
        f'<wp:anchor distT="0" distB="0" distL="0" distR="0"'
        f' simplePos="0" relativeHeight="{251660000 + sid}"'
        f' behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column">'
        f'<wp:posOffset>{to_emu(left_pt)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph">'
        f'<wp:posOffset>{to_emu(top_pt)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{to_emu(width_pt)}" cy="{to_emu(height_pt)}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapNone/>'
        f'<wp:docPr id="{sid}" name="Oval {sid}"/>'
        '<a:graphic>'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{to_emu(width_pt)}" cy="{to_emu(height_pt)}"/></a:xfrm>'
        '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
        '<a:noFill/>'
        '<a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>'
        '</wps:spPr><wps:bodyPr/></wps:wsp>'
        '</a:graphicData></a:graphic>'
        '</wp:anchor></w:drawing></w:r>'
    )
    para._element.append(etree.fromstring(xml))


def _circle_item_in_para(para, item_text, include_number=True, use_last=False):
    """段落テキスト内のitem_textを探し、赤い楕円シェイプを配置。
    テキストは変更せず、図形を上に重ねるだけ。
    use_last=Trueで最後の出現位置を使用（"有無"の"有"ではなく括弧内の"有"を選択）。
    """
    text = para.text
    idx = text.rfind(item_text) if use_last else text.find(item_text)
    if idx < 0:
        return

    target_start = idx
    target_end = idx + len(item_text)

    # 番号プレフィックスを含める
    if include_number and idx > 0:
        i = idx - 1
        while i >= 0 and text[i] in ' \u3000':
            i -= 1
        if i >= 0:
            j = i
            while j > 0 and text[j - 1] not in ' \u3000':
                j -= 1
            target_start = j

    font_size = _get_para_font_size_pt(para)
    half_w = font_size / 2  # 半角1文字 = font_size/2 pt

    # Word表セルの内部余白（デフォルト≈5.4pt）
    cell_margin = 5.4
    x_start = cell_margin + _display_width(text[:target_start]) * half_w
    item_w = _display_width(text[target_start:target_end]) * half_w

    height = font_size * 1.4
    top = -font_size * 0.2  # テキストの上下中央に配置

    _add_oval_to_para(para, x_start - 2, top, item_w + 4, height)


# ============================================================
# 段落テキスト操作（フォーマット保持）
# ============================================================
def _write_para_compact(para, text, max_chars=100):
    """テキスト長に応じてフォントサイズを自動縮小し行間も圧縮して書き込む。
    max_chars以下→9pt、1.5倍以下→8pt、それ以上→7ptを使用。"""
    text_len = len(str(text))
    if text_len <= max_chars:
        font_size = Pt(9)
    elif text_len <= max_chars * 1.5:
        font_size = Pt(8)
    else:
        font_size = Pt(7)
    _write_para(para, text, font_size_override=font_size)
    # 行間を圧縮（段落前後のスペースを0に）
    from docx.shared import Pt as _Pt
    para.paragraph_format.space_before = _Pt(0)
    para.paragraph_format.space_after = _Pt(0)


def _write_para(para, text, font_size_override=None):
    """段落のテキストを置換。既存ランのフォント情報を保持する。"""
    # 既存ランからフォント情報を取得
    font_name = None
    font_size = None
    if para.runs:
        r0 = para.runs[0]
        font_name = r0.font.name
        font_size = r0.font.size

    if font_size_override:
        font_size = font_size_override

    # 全ランをクリア
    for run in para.runs:
        run.text = ""
    # 既存ランがあれば最初のランに書き込み、なければ新規追加
    if para.runs:
        para.runs[0].text = str(text)
        if font_size_override:
            para.runs[0].font.size = font_size_override
    else:
        run = para.add_run(str(text))
        if font_name:
            run.font.name = font_name
        if font_size:
            run.font.size = font_size


def _append_to_para(para, text):
    """段落の末尾にテキストを追加。"""
    if para.runs:
        last_run = para.runs[-1]
        run = para.add_run(str(text))
        run.font.name = last_run.font.name
        run.font.size = last_run.font.size
    else:
        para.add_run(str(text))


# ============================================================
# 症状チェックリスト
# ============================================================
def _fill_symptoms(cell, cc):
    """R5C0の症状テキスト: 該当項目に赤い楕円シェイプを配置（テキスト非変更）"""
    categories = {
        "depressive_state": ([2], ["思考・運動抑制", "易刺激性・興奮", "憂うつ気分", "その他"]),
        "manic_state": ([4], ["行為心迫", "多弁", "感情高揚・易刺激性", "その他"]),
        "hallucination_delusion": ([6], ["幻覚", "妄想", "その他"]),
        "cognitive_decline": ([8], ["興奮", "昏迷", "拒絶", "その他"]),
        "personality_behavior": ([10], ["自閉", "感情平板化", "意欲の減退", "その他"]),
        "emotion_behavior": ([12, 13], [
            "爆発性", "易怒性", "気分変動", "暴力・衝動行為", "常同行為", "多動",
            "食行動の異常", "性行動の異常", "チック・汚言", "その他"]),
        "anxiety_neurosis": ([15, 16], [
            "強度の不安・恐怖感", "精神運動不穏", "心身衰弱", "強迫体験", "心気症状",
            "心的外傷に関連する症状", "解離・転換症状", "その他"]),
        "substance_use": ([23, 24], ["アルコール", "覚醒剤", "有機溶剤", "その他"]),
        "cognitive_learning": ([26, 27, 28, 29], ["知的障害", "認知症", "その他"]),
        "developmental": ([31, 32], [
            "相互的な社会関係の質的障害", "コミュニケーションのパターンにおける質的障害",
            "限定した常同的で反復的な関心と活動", "その他"]),
    }

    paras = cell.paragraphs

    for field, (para_idxs, items) in categories.items():
        selected = cc.get(field, [])
        if not selected:
            continue
        for pi in para_idxs:
            if pi < len(paras):
                for item in items:
                    if item in selected:
                        _circle_item_in_para(paras[pi], item)

    # (12) その他
    other = cc.get("other", "")
    if other and 33 < len(paras):
        _write_para(paras[33], f"（12）その他（{other}）")


def _fill_diagnosis_run(para, blank_run_idx, value):
    """病名行の空白run（全角スペース列）を値で置換する"""
    runs = para.runs
    if blank_run_idx < len(runs):
        runs[blank_run_idx].text = "　" + value + "　"


def _find_nested_table(table, row_idx, cell_idx):
    """テーブルのセル内にあるネストテーブルのXML要素を返す"""
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    try:
        row = table.rows[row_idx]
        tr = row._tr
        tcs = tr.findall(f'{ns}tc')
        if cell_idx < len(tcs):
            tc = tcs[cell_idx]
            nested = tc.find(f'{ns}tbl')
            return nested
    except (IndexError, AttributeError):
        pass
    return None


def _fill_icd_cells_xml(tbl_element, row_idx, code_str):
    """ネストテーブルのXML要素に1文字ずつICDコードを書き込む"""
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    rows = tbl_element.findall(f'{ns}tr')
    if row_idx >= len(rows):
        return
    tcs = rows[row_idx].findall(f'{ns}tc')
    for i, ch in enumerate(str(code_str)):
        if i >= len(tcs):
            break
        tc = tcs[i]
        # セル内の最初の段落の最初のrunにテキストを設定
        p = tc.find(f'{ns}p')
        if p is None:
            continue
        # 既存のrunを探すか新規作成
        r = p.find(f'{ns}r')
        if r is None:
            r = etree.SubElement(p, f'{ns}r')
        t = r.find(f'{ns}t')
        if t is None:
            t = etree.SubElement(r, f'{ns}t')
        t.text = ch


def _fill_label_value(cell, label_text, value, exact_para_idx=None):
    """セル内でラベルテキストを含む段落を見つけ、末尾に値を追記する"""
    if exact_para_idx is not None:
        para = cell.paragraphs[exact_para_idx]
        _write_para(para, value.strip() if not para.text.strip() else para.text)
        return

    for para in cell.paragraphs:
        if label_text in para.text:
            _append_to_para(para, value)
            return


def _fill_label_value_small(cell, label_text, value, font_size):
    """セル内でラベルテキストを含む段落を見つけ、小フォントで値を追記する"""
    for para in cell.paragraphs:
        if label_text in para.text:
            run = para.add_run(str(value))
            run.font.size = font_size
            return


def _fill_after_colon(cell, label_with_colon, value):
    """セル内で "ラベル：" を含む段落を見つけ、：の後の空白を値に置換する"""
    for para in cell.paragraphs:
        if label_with_colon in para.text:
            for run in para.runs:
                if label_with_colon in run.text:
                    # "ラベル：　　　" → "ラベル：値"
                    idx = run.text.find(label_with_colon) + len(label_with_colon)
                    run.text = run.text[:idx] + value + run.text[idx:].lstrip("　 ")
                    return
            # runが分割されている場合はラベル直後に追記
            _append_to_para(para, value)
            return


def _fill_designated_info_xml(doc, designated_num, years):
    """XMLレベルで「精神保健指定医の証の番号：」と「精神医療従事年数：」に値を埋める。
    元テンプレートの下線付き全角スペースの長さを維持する。"""
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    body = doc.element.body
    for t_elem in body.iter(f'{ns}t'):
        text = t_elem.text or ""
        if "精神保健指定医の証の番号：" in text and designated_num:
            orig_blank = "　　　　　　　　　"  # 元の全角スペース9文字
            # 値を入れた後、残りを全角スペースで埋めて下線の長さを維持
            filled = designated_num + "　" * max(0, len(orig_blank) - len(designated_num))
            t_elem.text = text.replace(
                "精神保健指定医の証の番号：" + orig_blank,
                "精神保健指定医の証の番号：" + filled
            )
            text = t_elem.text
        if "精神医療従事年数：" in text and years:
            orig_blank = "　　　　　　　　　　　　年"  # 元の全角スペース12文字+年
            filled = years + "　" * max(0, 12 - len(years)) + "年"
            t_elem.text = t_elem.text.replace(
                "精神医療従事年数：" + orig_blank,
                "精神医療従事年数：" + filled
            )


# ============================================================
# 福祉サービス
# ============================================================
def _fill_welfare(cell, wf):
    """福祉サービスの該当項目に赤い楕円シェイプを配置（テキスト非変更）"""
    items = wf.get("items", [])
    if not items:
        return
    other_detail = wf.get("other_detail", "")
    paras = cell.paragraphs

    service_map = {
        "自立訓練": 3,
        "共同生活援助": 4,
        "居宅介護": 5,
        "その他の障害福祉サービス": 6,
        "訪問指導等": 7,
        "なし": 8,
    }

    # その他の詳細を先に書き込む
    if other_detail and "その他の障害福祉サービス" in items and 6 < len(paras):
        text = paras[6].text
        text = text.replace("（　　　　　　　　　）", f"（{other_detail}）")
        _write_para(paras[6], text)

    for service_name, pi in service_map.items():
        if service_name in items and pi < len(paras):
            _circle_item_in_para(paras[pi], service_name)


# ============================================================
# ユーティリティ
# ============================================================
def _parse_ymd(d):
    if not d: return ("", "", "")
    parts = d.split("-")
    if len(parts) >= 3: return (parts[0], str(int(parts[1])), str(int(parts[2])))
    if len(parts) == 2: return (parts[0], str(int(parts[1])), "")
    return (d, "", "")

def _parse_ym(d):
    if not d: return ("", "")
    parts = d.split("-")
    if len(parts) >= 2: return (parts[0], str(int(parts[1])))
    return (d, "")

def _era_name(year_str):
    if not year_str: return ""
    try:
        y = int(year_str)
        if y >= 2019: return "令和"
        if y >= 1989: return "平成"
        if y >= 1926: return "昭和"
        if y >= 1912: return "大正"
        return "明治"
    except ValueError: return ""

def _era_year(year_str):
    if not year_str: return ""
    try:
        y = int(year_str)
        if y >= 2019: return str(y - 2018)
        if y >= 1989: return str(y - 1988)
        if y >= 1926: return str(y - 1925)
        if y >= 1912: return str(y - 1911)
        return str(y)
    except ValueError: return str(year_str)

def _format_postal_address(postal_code, address):
    postcode, body = _split_postal_code(postal_code or address)
    address_text = str(address).strip() if address else ""
    if postcode and not body:
        body = address_text.replace(f"〒{postcode}", "").replace(postcode, "").strip()
    if postcode and body: return f"〒{postcode} {body}"
    if postcode: return f"〒{postcode}"
    return body or address_text

def _split_postal_code(value):
    if not value: return ("", "")
    text = str(value).strip().replace("〒", "")
    if len(text) >= 8 and text[3] == "-" and text[:3].isdigit() and text[4:8].isdigit():
        return (text[:8], text[8:].strip())
    return ("", str(value).strip())

def _calc_age(birthdate, reference_date):
    by, bm, bd = _parse_ymd(birthdate)
    ry, rm, rd = _parse_ymd(reference_date)
    if not (by and bm and bd and ry and rm and rd): return ""
    try:
        birth_num = int(bm) * 100 + int(bd)
        ref_num = int(rm) * 100 + int(rd)
        age = int(ry) - int(by) - (1 if ref_num < birth_num else 0)
        return age if age >= 0 else ""
    except ValueError: return ""

def _merge_clinic_defaults(institution):
    merged = dict(institution or {})
    config_path = os.path.join(os.path.dirname(__file__), "..", "..", "config.json")
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            clinic = json.load(f).get("clinic", {})
    except Exception:
        clinic = {}
    for key in ("name", "address", "phone", "doctor", "code", "designated_number", "doctor_years"):
        if clinic.get(key):
            merged[key] = clinic[key]
    return merged


if __name__ == "__main__":
    main()
