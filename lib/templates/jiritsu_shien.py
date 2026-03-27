"""
自立支援医療診断書（精神通院医療用）テンプレート記入スクリプト
公式書式 2jiritsu.xls にJSONデータを流し込む（xlwings使用・フォーマット完全保持）
Usage: python jiritsu_shien.py <input.json> <output.xls>
"""
import sys, json, os, shutil, platform
import xlwings as xw

IS_MAC = platform.system() == "Darwin"

TEMPLATE = os.path.join(os.path.dirname(__file__), "..", "..", "sample", "自立支援", "2jiritsu.xls")

# ============================================================
# セル位置定義（すべて1-indexed）
# テンプレート 2jiritsu.xls の結合セル構造に基づく
# ============================================================
# 病名: 結合セル $AJ$26:$BB$28 → top-left (26, 36)
# ICDコード: ラベル $BD$27:$BL$28, 値は col 65 (BM27) 以降（非結合）
# 病歴ヘッダー: 結合 (40, 19)
# 病歴本文: 結合 $S$46:$CE$62 → top-left (46, 19)
# 症状チェック: 各行 col 4 から col 83 まで結合
# 条件詳細: 結合 $E$112:$CD$130 → top-left (112, 5)
# 治療内容: 結合 $D$132:$CE$162 → top-left (132, 4)
# 治療方針: 結合 $E$168:$CD$177 → top-left (168, 5)
# 福祉サービス: 各行 col 5
# 備考: 結合 (179, 48) 以降
# 日付: 結合 $I$193:$AR$195 → top-left (193, 9)
# 医療機関情報: 値は結合 $S$199:$AR$200 等 → top-left col 19
# ============================================================


def main():
    if len(sys.argv) < 3:
        print("Usage: python jiritsu_shien.py <input.json> <output.xls>", file=sys.stderr)
        sys.exit(1)
    input_path, output_path = sys.argv[1], sys.argv[2]
    data = json.loads(open(input_path, "r", encoding="utf-8").read())
    fill_template(data, output_path)
    print(f"生成完了: {output_path}")


def fill_template(data, output_path):
    abs_output = os.path.abspath(output_path)
    shutil.copy2(TEMPLATE, abs_output)

    app = xw.App(visible=False)
    try:
        wb = app.books.open(abs_output)
        ws = wb.sheets[0]

        ins = data.get("institution", {})
        ins = _merge_clinic_defaults(ins)
        cc = data.get("current_condition", {})
        treat = data.get("treatment", {})
        wf = data.get("welfare_services", {})

        _fill_patient_info(ws, data.get("patient", {}), data.get("date", ""))

        # ---- 病名 ----
        # (1) 主たる精神障害: 結合セル (26,36)-(28,54)
        if data.get("diagnosis_main"):
            _w(ws, 26, 36, data["diagnosis_main"])
        # ICDコードラベルの位置補正（中央揃え）
        ws.range((27, 56)).api.HorizontalAlignment = -4108  # xlCenter
        # ICDコードはラベル右側の記入欄に書く
        if data.get("icd_code_main"):
            _write_boxed_text(ws, 27, 65, data["icd_code_main"])
        # (2) 従たる精神障害
        if data.get("diagnosis_sub"):
            _w(ws, 29, 36, data["diagnosis_sub"])
        if data.get("icd_code_sub"):
            _write_boxed_text(ws, 30, 65, data["icd_code_sub"])
        # (3) 身体合併症
        if data.get("comorbidity"):
            _w(ws, 32, 36, data["comorbidity"])

        # ---- ２ 発病から現在までの病歴 ----
        onset = data.get("onset_date", "")
        fv = data.get("first_visit_date", "")
        onset_y, onset_m = _parse_ym(onset)
        fv_y, fv_m, fv_d = _parse_ymd(fv)
        header = f"（推定発病時期　{onset_y or '　　'}年{onset_m or '　　'}月頃、　初診日　{fv_y or '　　'}年{fv_m or '　'}月{fv_d or '　'}日）"
        _w(ws, 40, 19, header)

        course = data.get("clinical_course", "")
        if course:
            _w(ws, 46, 19, course)

        # ---- ３ 現在の病状、状態像等 ----
        _fill_symptoms(ws, cc)

        # ---- ４ 病状の具体的程度 ----
        # 結合セル $E$112:$CD$130 → top-left (112, 5)
        detail = data.get("condition_detail", "")
        if detail:
            _w(ws, 112, 5, detail)

        # ---- ５ 現在の治療内容 ----
        # 結合セル $D$132:$CE$162 → top-left (132, 4)
        # テンプレートのフォーマットに合わせて書き込む
        meds = treat.get("medications", [])
        med_text = "\n".join(f"・{m['name']} {m.get('dosage','')} {m.get('frequency','')}" for m in meds)
        psych = treat.get("psychotherapy", "")
        designated = treat.get("designated_doctor", "")

        # 訪問看護指示: テキストは中立のまま書き込み、図形の○で選択を示す
        nursing_text = "（３）訪問看護指示の有無（　有　・　無　）"

        treat_text = (
            " ５　現在の治療内容\n"
            "（１）投薬内容　（※診断名に対する薬剤名（商品名可）をご記載ください。）\n"
            f"{med_text}\n\n"
            "（２）精神療法等\n\n\n\n"
            f"{psych}\n\n"
            f"{nursing_text}"
        )
        _w(ws, 132, 4, treat_text)

        # 有/無 に図形の○を配置
        nursing_target = "有" if designated == "有" else "無"
        _circle_in_treatment_cell(ws, 132, 4, treat_text, nursing_target)

        # ---- ６ 今後の治療方針 ----
        # 結合セル $E$168:$CD$177 → top-left (168, 5)
        plan = data.get("treatment_plan", "")
        if plan:
            _w(ws, 168, 5, plan)

        # ---- ７ 障害福祉サービス ----
        ws_items = wf.get("items", [])
        _fill_welfare(ws, ws_items, wf.get("other_detail", ""))

        # ---- ８ 備考 ----
        remarks = data.get("remarks", "")
        if remarks:
            _w(ws, 179, 59, remarks)

        # ---- 日付・医療機関 ----
        d = data.get("date", "")
        d_y, d_m, d_d = _parse_ymd(d)
        date_str = f"{d_y or '　　'}年　{d_m or '　'}月　{d_d or '　'}日　医療機関コード"
        code = ins.get("code", "")
        if code:
            date_str += f"　{code}"
        _w(ws, 193, 9, date_str)

        # 医療機関情報: 結合セル $S$N:$AR$N+1
        if ins.get("address"):
            _w(ws, 199, 19, ins["address"])
        if ins.get("name"):
            _w(ws, 202, 19, ins["name"])
        if ins.get("phone"):
            _w(ws, 205, 19, ins["phone"])
        if ins.get("doctor"):
            _write_footer_doctor(ws, ins["doctor"])

        # 精神保健指定医の証の番号
        designated_num = ins.get("designated_number", "")
        if designated_num:
            _w(ws, 210, 46, f"精神保健指定医の証の番号：{designated_num}")

        # 精神保健従事年数
        years = ins.get("doctor_years", "")
        if years and str(years).strip():
            _w(ws, 213, 46, f"精神保健従事年数：{years}年")

        wb.save()
        wb.close()
    finally:
        app.quit()


def _w(ws, row, col, value):
    """Write value to cell (1-indexed row, col) — preserves existing cell formatting"""
    ws.range((row, col)).value = value


def _write_boxed_text(ws, row, start_col, value, max_chars=12, start_offset=2):
    text = str(value or "")
    for offset in range(max_chars):
        ws.range((row, start_col + offset)).value = ""
    cell = ws.range((row, start_col + start_offset))
    cell.value = text[:max_chars]
    cell.api.Font.Size = 8
    cell.api.VerticalAlignment = -4108  # xlCenter


def _write_footer_doctor(ws, doctor):
    target = ws.range((208, 19), (209, 44))
    if not target.api.MergeCells:
        target.api.Merge()
    target.value = doctor


# ============================================================
# 図形の○（シェイプ）ヘルパー — テキスト非変更で選択肢を囲む
# フォントサイズから文字幅を直接計算して位置を決定する
# ============================================================

# セル左端のインデント（pt）
_CELL_MARGIN = 2


def _add_circle_shape(ws, left, top, width, height):
    """透明な楕円図形（○マーク）を追加。赤色で視認性を確保。"""
    shape = ws.api.Shapes.AddShape(9, left, top, width, height)  # 9 = msoShapeOval
    shape.Fill.Visible = 0  # 透明
    shape.Line.ForeColor.RGB = 255  # 赤
    shape.Line.Weight = 1.5
    return shape


def _get_font_size(ws, row, col):
    """セルのフォントサイズ（pt）を取得"""
    try:
        size = ws.range((row, col)).api.Font.Size
        return size if size and size > 0 else 11
    except Exception:
        return 11


def _text_width_pt(text, char_w):
    """テキストのポイント幅を推定（全角=char_w, 半角=char_w/2）"""
    w = 0.0
    for ch in text:
        cp = ord(ch)
        if cp <= 0x7E or (0xFF61 <= cp <= 0xFF9F):
            w += char_w * 0.5
        else:
            w += char_w
    return w


def _circle_item_in_cell(ws, row, col, item_text, include_number=True):
    """結合セル内のテキストを探し、番号ごと図形の○で囲む。

    フォントサイズから文字幅を計算し、セル左端からの絶対位置で配置。
    テンプレート文字列を変更せず、図形を上に重ねるだけなので
    Excel上で図形を移動・削除するだけで修正できる。
    """
    cell = ws.range((row, col))
    cell_text = str(cell.value or "")
    idx = cell_text.find(item_text)
    if idx < 0:
        return

    target_start = idx
    target_end = idx + len(item_text)

    # 番号プレフィックスを含める（例: "（１）思考..." → "（１）" 部分も囲む）
    if include_number and idx > 0:
        i = idx - 1
        while i >= 0 and cell_text[i] in ' \u3000':
            i -= 1
        if i >= 0:
            j = i
            while j > 0 and cell_text[j - 1] not in ' \u3000':
                j -= 1
            target_start = j

    ma = cell.api.MergeArea
    cell_left = ma.Left
    cell_top = ma.Top
    cell_height = ma.Height

    char_w = _get_font_size(ws, row, col)
    x_offset = _text_width_pt(cell_text[:target_start], char_w)
    x_width = _text_width_pt(cell_text[target_start:target_end], char_w)

    pad_x, pad_y = 3, 2
    _add_circle_shape(
        ws,
        cell_left + _CELL_MARGIN + x_offset - pad_x,
        cell_top - pad_y,
        x_width + 2 * pad_x,
        cell_height + 2 * pad_y
    )


def _circle_in_treatment_cell(ws, row, col, full_text, target):
    """治療内容セル（複数行）内の「有」または「無」に図形の○を配置。

    改行で行を分割し、フォントサイズから垂直・水平位置を計算して配置する。
    """
    lines = full_text.split('\n')
    target_line_idx = None
    target_line = ""
    for i, line in enumerate(lines):
        if target in line:
            target_line_idx = i
            target_line = line

    if target_line_idx is None:
        return

    cell = ws.range((row, col))
    ma = cell.api.MergeArea
    cell_left = ma.Left
    cell_top = ma.Top

    char_w = _get_font_size(ws, row, col)

    # 「　有　」「　無　」のように全角スペースで囲まれた文字を探す
    search = f"\u3000{target}\u3000"
    search_idx = target_line.find(search)
    if search_idx >= 0:
        char_idx = search_idx + 1
    else:
        char_idx = target_line.rfind(target)
    if char_idx < 0:
        return

    x_offset = _text_width_pt(target_line[:char_idx], char_w)
    x_width = _text_width_pt(target, char_w)

    line_height = char_w * 1.3
    y = cell_top + target_line_idx * line_height
    h = line_height

    pad = 5
    _add_circle_shape(
        ws,
        cell_left + _CELL_MARGIN + x_offset - pad,
        y,
        x_width + 2 * pad,
        h + 2 * pad
    )


def _fill_patient_info(ws, patient, reference_date):
    furigana = (
        patient.get("furigana")
        or patient.get("name_kana")
        or patient.get("kana")
        or ""
    )
    if furigana:
        _w(ws, 13, 16, furigana)

    if patient.get("name"):
        _w(ws, 15, 16, patient["name"])

    birthdate = patient.get("birthdate", "")
    age = patient.get("age")
    if age in (None, "") and birthdate:
        age = _calc_age(birthdate, reference_date)
    birth_text = _format_birthdate_text(birthdate, age)
    if birth_text:
        _w(ws, 16, 37, birth_text)

    postal_code = patient.get("postal_code", "") or patient.get("zip_code", "")
    address = patient.get("address", "")
    address_text = _format_postal_address(postal_code, address)
    if address_text:
        _w(ws, 20, 16, address_text)


def _fill_symptoms(ws, cc):
    """症状チェックリスト: 該当項目に図形の○を配置（テンプレートテキスト非変更）"""
    symptom_rows = {
        66: ("depressive_state", ["思考・運動抑制", "易刺激性・興奮", "憂うつ気分", "その他"]),
        68: ("manic_state", ["行為心迫", "多弁", "感情高揚・易刺激性", "その他"]),
        70: ("hallucination_delusion", ["幻覚", "妄想", "その他"]),
        72: ("psychomotor_excitement", ["興奮", "昏迷", "拒絶", "その他"]),
        74: ("residual_state", ["自閉", "感情平板化", "意欲の減退", "その他"]),
    }

    emotion_items_r76 = ["爆発性", "易怒性", "気分変動", "暴力・衝動行為", "常同行為", "多動"]
    emotion_items_r77 = ["食行動の異常", "性行動の異常", "チック・汚言", "その他"]

    anxiety_items_r79 = ["強度の不安・恐怖感", "精神運動不穏", "心身衰弱", "強迫体験", "心気症状"]
    anxiety_items_r80 = ["心的外傷に関連する症状", "解離・転換症状", "その他"]

    for row, (field, items) in symptom_rows.items():
        selected = cc.get(field, [])
        if not selected and field == "psychomotor_excitement":
            selected = cc.get("cognitive_decline", [])
        if not selected and field == "residual_state":
            selected = cc.get("personality_behavior", [])
        if selected:
            for item in items:
                if item in selected:
                    _circle_item_in_cell(ws, row, 4, item)

    emo = cc.get("emotion_behavior", [])
    if emo:
        for item in emotion_items_r76:
            if item in emo:
                _circle_item_in_cell(ws, 76, 4, item)
        for item in emotion_items_r77:
            if item in emo:
                _circle_item_in_cell(ws, 77, 4, item)

    anx = cc.get("anxiety_neurosis", [])
    if anx:
        for item in anxiety_items_r79:
            if item in anx:
                _circle_item_in_cell(ws, 79, 4, item)
        for item in anxiety_items_r80:
            if item in anx:
                _circle_item_in_cell(ws, 80, 4, item)

    epi = cc.get("epilepsy", {})
    if epi.get("has_epilepsy"):
        stype = epi.get("seizure_type", "")
        freq = epi.get("frequency", "")
        last = epi.get("last_seizure", "")
        last_y, last_m, last_d = _parse_ymd(last)
        new_text = f"　　１てんかん発作　発作型（{stype if stype else 'イ・ロ・ハ・ニ'}） 頻度（{freq or '　　'}回／月・年） 最終発作（{last_y or '　　'}年{last_m or '　'}月{last_d or '　'}日）"
        _w(ws, 83, 4, new_text)
        _circle_item_in_cell(ws, 83, 4, "１てんかん発作")
        if stype:
            _circle_item_in_cell(ws, 83, 4, stype, include_number=False)

    sub = cc.get("substance_use", [])
    if sub:
        for item in ["アルコール", "覚醒剤", "有機溶剤", "その他"]:
            if item in sub:
                _circle_item_in_cell(ws, 88, 4, item)

    cog = cc.get("cognitive_learning", [])
    if cog:
        for item in ["知的障害"]:
            if item in cog:
                _circle_item_in_cell(ws, 91, 4, item)

    dev = cc.get("developmental", [])
    if dev:
        for item in ["相互的な社会関係の質的障害", "コミュニケーションのパターンにおける質的障害"]:
            if item in dev:
                _circle_item_in_cell(ws, 96, 4, item)
        for item in ["限定した常同的で反復的な関心と活動", "その他"]:
            if item in dev:
                _circle_item_in_cell(ws, 97, 4, item)


def _fill_welfare(ws, items, other_detail):
    """福祉サービス: 該当項目に図形の○を配置（テンプレートテキスト非変更）"""
    mapping = {
        "自立訓練": 183,
        "共同生活援助": 184,
        "居宅介護": 185,
        "その他の障害福祉サービス": 186,
        "訪問指導等": 187,
        "なし": 188,
    }
    # その他の詳細を先に書き込む（テキスト変更後に図形配置）
    if other_detail and "その他の障害福祉サービス" in items:
        orig = ws.range((186, 5)).value or ""
        _w(ws, 186, 5, f"{str(orig).rstrip()}（{other_detail}）" if orig else f"その他の障害福祉サービス（{other_detail}）")

    for item_name, row in mapping.items():
        if item_name in items:
            _circle_item_in_cell(ws, row, 5, item_name)


def _parse_ymd(d):
    if not d:
        return ("", "", "")
    parts = d.split("-")
    if len(parts) >= 3:
        return (parts[0], str(int(parts[1])), str(int(parts[2])))
    if len(parts) == 2:
        return (parts[0], str(int(parts[1])), "")
    return (d, "", "")

def _parse_ym(d):
    if not d:
        return ("", "")
    parts = d.split("-")
    if len(parts) >= 2:
        return (parts[0], str(int(parts[1])))
    return (d, "")



def _era_name(year_str):
    if not year_str:
        return ""
    try:
        y = int(year_str)
        if y >= 2019:
            return "令和"
        if y >= 1989:
            return "平成"
        if y >= 1926:
            return "昭和"
        if y >= 1912:
            return "大正"
        return "明治"
    except ValueError:
        return ""


def _era_year(year_str):
    if not year_str:
        return ""
    try:
        y = int(year_str)
        if y >= 2019:
            return str(y - 2018)
        if y >= 1989:
            return str(y - 1988)
        if y >= 1926:
            return str(y - 1925)
        if y >= 1912:
            return str(y - 1911)
        if y >= 1868:
            return str(y - 1867)
        return str(y)
    except ValueError:
        return str(year_str)


def _format_birthdate_text(birthdate, age):
    if not birthdate:
        return ""
    y, m, d = _parse_ymd(birthdate)
    parts = []
    era = _era_name(y)
    era_year = _era_year(y) if y else ""
    if era:
        parts.append(era)
    if era_year:
        parts.append(f"{era_year}年")
    if m:
        parts.append(f"{m}月")
    if d:
        parts.append(f"{d}日生")
    if age not in (None, ""):
        parts.append(f"（{age}歳）")
    return "".join(parts)


def _format_postal_address(postal_code, address):
    postcode, body = _split_postal_code(postal_code or address)
    address_text = str(address).strip() if address else ""
    if postcode and not body:
        body = address_text.replace(f"〒{postcode}", "").replace(postcode, "").strip()
    if postcode and body:
        return f"〒{postcode} {body}"
    if postcode:
        return f"〒{postcode}"
    return body or address_text


def _split_postal_code(value):
    if not value:
        return ("", "")
    text = str(value).strip().replace("〒", "")
    if len(text) >= 8 and text[3] == "-" and text[:3].isdigit() and text[4:8].isdigit():
        return (text[:8], text[8:].strip())
    return ("", str(value).strip())


def _calc_age(birthdate, reference_date):
    by, bm, bd = _parse_ymd(birthdate)
    ry, rm, rd = _parse_ymd(reference_date)
    if not (by and bm and bd and ry and rm and rd):
        return ""
    try:
        birth_num = int(bm) * 100 + int(bd)
        ref_num = int(rm) * 100 + int(rd)
        age = int(ry) - int(by) - (1 if ref_num < birth_num else 0)
        return age if age >= 0 else ""
    except ValueError:
        return ""


def _merge_clinic_defaults(institution):
    merged = dict(institution or {})
    config_path = os.path.join(os.path.dirname(__file__), "..", "..", "config.json")
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            clinic = json.load(f).get("clinic", {})
    except Exception:
        clinic = {}
    field_map = {
        "name": "name",
        "address": "address",
        "phone": "phone",
        "doctor": "doctor",
        "code": "code",
        "designated_number": "designated_number",
        "doctor_years": "doctor_years",
    }
    for dst, src in field_map.items():
        if clinic.get(src):
            merged[dst] = clinic[src]
    return merged


if __name__ == "__main__":
    main()
