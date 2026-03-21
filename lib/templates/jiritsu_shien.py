"""
自立支援医療診断書（精神通院医療用）テンプレート記入スクリプト
公式書式 2jiritsu.xls にJSONデータを流し込む（xlwings使用・フォーマット完全保持）
Usage: python jiritsu_shien.py <input.json> <output.xls>
"""
import sys, json, os, shutil
import xlwings as xw

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

        # 訪問看護指示: テンプレートは「有　・　無」形式。選択側に○をつける
        if designated == "有":
            nursing_text = "（３）訪問看護指示の有無（　○有　・　無　）"
        else:
            nursing_text = "（３）訪問看護指示の有無（　有　・　○無　）"

        treat_text = (
            " ５　現在の治療内容\n"
            "（１）投薬内容　（※診断名に対する薬剤名（商品名可）をご記載ください。）\n"
            f"{med_text}\n\n"
            "（２）精神療法等\n\n\n\n"
            f"{psych}\n\n"
            f"{nursing_text}"
        )
        _w(ws, 132, 4, treat_text)

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
    """症状チェックリスト: 該当項目番号の前に○を追加"""
    symptom_rows = {
        # 1-indexed row, field_name, items_in_order
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
            orig = ws.range((row, 4)).value or ""
            new_text = _mark_items_in_text(orig, items, selected)
            _w(ws, row, 4, new_text)

    emo = cc.get("emotion_behavior", [])
    if emo:
        orig76 = ws.range((76, 4)).value or ""
        _w(ws, 76, 4, _mark_items_in_text(orig76, emotion_items_r76, emo))
        orig77 = ws.range((77, 4)).value or ""
        _w(ws, 77, 4, _mark_items_in_text(orig77, emotion_items_r77, emo))

    anx = cc.get("anxiety_neurosis", [])
    if anx:
        orig79 = ws.range((79, 4)).value or ""
        _w(ws, 79, 4, _mark_items_in_text(orig79, anxiety_items_r79, anx))
        orig80 = ws.range((80, 4)).value or ""
        _w(ws, 80, 4, _mark_items_in_text(orig80, anxiety_items_r80, anx))

    epi = cc.get("epilepsy", {})
    if epi.get("has_epilepsy"):
        stype = epi.get("seizure_type", "")
        freq = epi.get("frequency", "")
        last = epi.get("last_seizure", "")
        last_y, last_m, last_d = _parse_ymd(last)
        new_text = f"　　○１てんかん発作　発作型（{'○' + stype if stype else 'イ・ロ・ハ・ニ'}） 頻度（{freq or '　　'}回／月・年） 最終発作（{last_y or '　　'}年{last_m or '　'}月{last_d or '　'}日）"
        _w(ws, 83, 4, new_text)

    sub = cc.get("substance_use", [])
    if sub:
        orig88 = ws.range((88, 4)).value or ""
        _w(ws, 88, 4, _mark_items_in_text(orig88, ["アルコール", "覚醒剤", "有機溶剤", "その他"], sub))

    cog = cc.get("cognitive_learning", [])
    if cog:
        orig91 = ws.range((91, 4)).value or ""
        _w(ws, 91, 4, _mark_items_in_text(orig91, ["知的障害"], cog))

    dev = cc.get("developmental", [])
    if dev:
        orig96 = ws.range((96, 4)).value or ""
        _w(ws, 96, 4, _mark_items_in_text(orig96, ["相互的な社会関係の質的障害", "コミュニケーションのパターンにおける質的障害"], dev))
        orig97 = ws.range((97, 4)).value or ""
        _w(ws, 97, 4, _mark_items_in_text(orig97, ["限定した常同的で反復的な関心と活動", "その他"], dev))


def _mark_items_in_text(original_text, item_names, selected):
    """テキスト中の該当する項目番号を丸数字に置き換える"""
    if not original_text or not selected:
        return original_text
    text = str(original_text)
    for item_name in item_names:
        if item_name in selected:
            idx = text.find(item_name)
            if idx > 0:
                text = _circle_nearest_number(text, idx)
    return text


def _fill_welfare(ws, items, other_detail):
    """福祉サービス: 該当項目に○を追加"""
    mapping = {
        "自立訓練": 183,
        "共同生活援助": 184,
        "居宅介護": 185,
        "その他の障害福祉サービス": 186,
        "訪問指導等": 187,
        "なし": 188,
    }
    for item_name, row in mapping.items():
        if item_name in items:
            orig = ws.range((row, 5)).value or ""
            if orig:
                _w(ws, row, 5, _circle_first_number(str(orig).lstrip()))

    if other_detail and "その他の障害福祉サービス" in items:
        _w(ws, 186, 5, f"④その他の障害福祉サービス（{other_detail}）")


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


def _circle_nearest_number(text, idx):
    number_map = {
        "1": "①", "2": "②", "3": "③", "4": "④", "5": "⑤",
        "6": "⑥", "7": "⑦", "8": "⑧", "9": "⑨", "10": "⑩",
        "１": "①", "２": "②", "３": "③", "４": "④", "５": "⑤",
        "６": "⑥", "７": "⑦", "８": "⑧", "９": "⑨", "１０": "⑩",
    }
    i = idx - 1
    while i >= 0 and text[i] in " \u3000":
        i -= 1
    if i < 0:
        return text

    if text[i] in "）)":
        end = i
        start = i - 1
        while start >= 0 and text[start] not in "（(":
            start -= 1
        if start >= 0:
            token = text[start + 1:end]
            circled = number_map.get(token)
            if circled:
                return text[:start] + circled + text[end + 1:]

    end = i
    start = i
    while start >= 0 and text[start] in "0123456789０１２３４５６７８９":
        start -= 1
    token = text[start + 1:end + 1]
    circled = number_map.get(token)
    if circled:
        return text[:start + 1] + circled + text[end + 1:]
    return text


def _circle_first_number(text):
    for idx, ch in enumerate(text):
        if ch in "0123456789０１２３４５６７８９（(":
            return _circle_nearest_number(text, idx + 2)
    return text


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
