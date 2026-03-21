"""
紹介状（診療情報提供書）テキスト版生成スクリプト
JSONデータから紹介状の文面のみをプレーンテキストで出力する
手動でフォーマットにコピペして使用する

Usage: python referral_txt.py <input.json> <output.txt>
"""
import sys, json, os
from datetime import datetime


def main():
    if len(sys.argv) < 3:
        print("Usage: python referral_txt.py <input.json> <output.txt>", file=sys.stderr)
        sys.exit(1)
    data = json.loads(open(sys.argv[1], "r", encoding="utf-8").read())
    text = build_text(data)
    with open(sys.argv[2], "w", encoding="utf-8") as f:
        f.write(text)
    print(f"生成完了: {sys.argv[2]}")


def build_text(data):
    ins = _merge_clinic_defaults(data)
    patient = data.get("patient", {})

    lines = []

    # 日付
    date_str = _format_date(data.get("date", ""))
    lines.append(f"                                        {date_str}")
    lines.append("")

    # 宛先
    to_inst = data.get("to_institution", "")
    to_dept = data.get("to_department", "")
    to_doc = data.get("to_doctor", "御担当医")
    lines.append(f"{to_inst} {to_dept}")
    lines.append(f"{to_doc} 先生　御侍史")
    lines.append("")

    # 差出人
    lines.append(f"                        {ins.get('name', '')}")
    lines.append(f"                        {ins.get('department', '')}")
    lines.append(f"                        {ins.get('address', '')}")
    lines.append(f"                        TEL: {ins.get('phone', '')}")
    lines.append(f"                        医師　{ins.get('doctor', '')}")
    lines.append("")
    lines.append("=" * 60)

    # 患者情報
    lines.append("")
    lines.append(f"【患者氏名】{patient.get('name', '')}")
    bd = _format_date(patient.get("birthdate", ""))
    age = patient.get("age", "")
    sex = patient.get("sex", "")
    lines.append(f"【生年月日】{bd}（{age}歳）　【性別】{sex}")
    lines.append("")

    # 傷病名
    diag = data.get("diagnosis", [])
    if diag:
        lines.append("【傷病名】")
        for d in diag:
            lines.append(f"　{d}")
        lines.append("")

    # 症状経過および治療経過
    course = data.get("clinical_course", "")
    if course:
        lines.append("【症状経過および治療経過】")
        lines.append(course)
        lines.append("")

    # 現在の処方
    meds = data.get("current_medications", [])
    if meds:
        lines.append("【現在の処方】")
        for m in meds:
            name = m.get("name", "")
            dosage = m.get("dosage", "")
            freq = m.get("frequency", "")
            lines.append(f"　{name} {dosage} {freq}")
        lines.append("")

    # 検査結果
    tests = data.get("test_results", "")
    if tests:
        lines.append("【検査結果】")
        lines.append(tests)
        lines.append("")

    # 紹介目的
    purpose = data.get("purpose", "")
    if purpose:
        lines.append("【紹介目的】")
        lines.append(purpose)
        lines.append("")

    # 備考
    remarks = data.get("remarks", "")
    if remarks:
        lines.append("【備考】")
        lines.append(remarks)
        lines.append("")

    return "\n".join(lines)


def _format_date(date_str):
    if not date_str:
        return ""
    parts = date_str.split("-")
    if len(parts) >= 3:
        return f"{parts[0]}年{int(parts[1])}月{int(parts[2])}日"
    if len(parts) == 2:
        return f"{parts[0]}年{int(parts[1])}月"
    return date_str


def _merge_clinic_defaults(data):
    """config.jsonのクリニック情報をマージして返す"""
    merged = {}
    # JSONの from_* フィールド
    for key, field in [("name", "from_institution"), ("department", "from_department"),
                       ("doctor", "from_doctor"), ("address", "from_address"),
                       ("phone", "from_phone")]:
        if data.get(field):
            merged[key] = data[field]

    # config.json から補完
    config_path = os.path.join(os.path.dirname(__file__), "..", "..", "config.json")
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            clinic = json.load(f).get("clinic", {})
    except Exception:
        clinic = {}
    for key in ("name", "department", "doctor", "address", "phone"):
        if not merged.get(key) and clinic.get(key):
            merged[key] = clinic[key]

    return merged


if __name__ == "__main__":
    main()
