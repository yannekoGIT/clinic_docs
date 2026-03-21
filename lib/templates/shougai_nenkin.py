"""
障害年金診断書（精神の障害用）様式第120号の4 テンプレート記入スクリプト
公式書式 04 (1).xlsx にJSONデータを流し込む（xlwings使用・フォーマット完全保持）
Usage: python shougai_nenkin.py <input.json> <output.xlsx>
"""
import sys, json, os, shutil, platform
import xlwings as xw

IS_MAC = platform.system() == "Darwin"

TEMPLATE = os.path.join(os.path.dirname(__file__), "..", "..", "sample", "障害年金", "04 (1).xlsx")
C = "レ"  # チェックマーク（プルダウン "レ,　" の選択値）

def main():
    if len(sys.argv) < 3:
        print("Usage: python shougai_nenkin.py <input.json> <output.xlsx>", file=sys.stderr)
        sys.exit(1)
    data = json.loads(open(sys.argv[1], "r", encoding="utf-8").read())
    fill_template(data, sys.argv[2])
    print(f"生成完了: {sys.argv[2]}")

def s(ws, cell_ref, value, shrink=False):
    """セルに値を書き込む（xlwings経由。フォーマット完全保持）
    shrink=True: 枠に合わせてフォントサイズを自動縮小（縮小して全体を表示）
    """
    cell = ws.range(cell_ref)
    # リテラル "\n" がテキストに含まれる場合、実際の改行に変換
    if isinstance(value, str):
        value = value.replace("\\n", "\n")
    cell.value = value
    if shrink:
        try:
            cell.api.ShrinkToFit = True
        except Exception:
            pass  # Mac版Excelでは未対応の場合あり

def fill_template(data, output_path):
    abs_output = os.path.abspath(output_path)
    shutil.copy2(TEMPLATE, abs_output)

    app = xw.App(visible=False)
    try:
        wb = app.books.open(abs_output)
        ws = wb.sheets[0]

        # シート保護を解除（パスワードなし）
        try:
            ws.api.Unprotect()
        except Exception:
            pass

        pt = data.get("patient", {})
        diag = data.get("diagnosis", {})
        fv = data.get("first_visit", {})
        rem = data.get("remission", {})
        hist = data.get("history", {})
        init = data.get("initial_findings", {})
        dev = data.get("development_history", {})
        ds = data.get("disability_state", {})
        sc = ds.get("symptom_checklist", {})
        dl = ds.get("daily_living", {})
        da = dl.get("daily_assessment", {})
        emp = ds.get("employment", {})
        ins = data.get("institution", {})
        ins = _merge_clinic_defaults(ins)

        # ======== ヘッダー: 「精」に○（図形で囲む） ========
        try:
            _add_circle_around_cell(ws, "H3")
        except Exception:
            pass  # Mac版で図形APIが異なる場合はスキップ

        # ======== 患者基本情報 ========
        _fill_patient_info(ws, pt, data.get("date", ""))

        # ======== ① 傷病名 ========
        if diag.get("disease_name"):
            s(ws, "O20", diag["disease_name"])
        if diag.get("icd_code"):
            s(ws, "AA26", diag["icd_code"])

        # ======== ② 発生年月日 ========
        onset = diag.get("onset_date", "")
        oy, om, od = _ymd(onset)
        if diag.get("onset_source") in ("診療録で確認", "診療録で確定"):
            s(ws, "BS20", C)
        elif diag.get("onset_source") == "本人の申立て":
            s(ws, "BS21", C)
        era = _era(oy)
        if era:
            s(ws, "AX20", era)
        if oy: s(ws, "BC20", _era_year(oy))
        if om: s(ws, "BH20", om)
        if od: s(ws, "BM20", od)
        if diag.get("occupation_at_onset"):
            s(ws, "CO20", diag["occupation_at_onset"], shrink=True)

        # ======== ③ 初診日 ========
        fv_date = fv.get("date", "")
        fy, fm, fd_val = _ymd(fv_date)
        if fv.get("source") in ("診療録で確認", "診療録で確定"):
            s(ws, "BS25", C)
        elif fv.get("source") == "本人の申立て":
            s(ws, "BS26", C)
        era_fv = _era(fy)
        if era_fv:
            s(ws, "AX25", era_fv)
        if fy: s(ws, "BC25", _era_year(fy))
        if fm: s(ws, "BH25", fm)
        if fd_val: s(ws, "BM25", fd_val)

        # ======== ④ 既存障害 ========
        if data.get("existing_disability"):
            s(ws, "CO25", data["existing_disability"])

        # ======== ⑤ 既往症 ========
        if data.get("prior_illness"):
            s(ws, "CO30", data["prior_illness"])

        # ======== ⑥ 傷病が治ったかどうか ========
        if rem.get("remission_date"):
            ry, rm_val, rd = _ymd(rem["remission_date"])
            if ry: s(ws, "AA30", str(ry))
            if rm_val: s(ws, "AF30", str(rm_val))
            if rd: s(ws, "AK30", str(rd))
            if rem.get("confirmed_or_estimated") == "確認":
                s(ws, "AQ30", C)
            elif rem.get("confirmed_or_estimated") == "推定":
                s(ws, "AQ32", C)
        outlook = rem.get("prognosis_outlook", "")
        if outlook == "有": s(ws, "BL31", C)
        elif outlook == "無": s(ws, "BR31", C)
        elif outlook == "不明": s(ws, "BX31", C)

        # ======== ⑦ 病歴 ========
        if hist.get("informant_name"):
            s(ws, "AI36", hist["informant_name"])
        if hist.get("informant_relationship"):
            s(ws, "BK36", hist["informant_relationship"])
        if hist.get("interview_date"):
            iy, im, id_val = _ymd(hist["interview_date"])
            if iy: s(ws, "CI36", _era_year(iy))
            if im: s(ws, "CQ36", im)
            if id_val: s(ws, "CW36", id_val)
        if hist.get("narrative"):
            s(ws, "V37", hist["narrative"])

        # ======== ⑧ 初診時所見 ========
        if init.get("first_visit_date"):
            iy8, im8, id8 = _ymd(init["first_visit_date"])
            era8 = _era(iy8)
            if era8: s(ws, "C47", era8)
            if iy8: s(ws, "H47", _era_year(iy8))
            if im8: s(ws, "L47", im8)
            if id8: s(ws, "P47", id8)
        if init.get("findings"):
            s(ws, "V44", init["findings"])

        # ======== ⑨ 発育・養育歴等 ========
        if dev.get("development"):
            s(ws, "V52", dev["development"])
        edu = dev.get("education", {})
        if edu:
            _fill_education(ws, edu)
        if dev.get("work_history"):
            s(ws, "CH52", dev["work_history"])

        # エ 治療歴
        treat_hist = dev.get("treatment_history", [])
        _fill_treatment_history(ws, treat_hist)

        # ======== ⑩ 障害の状態 ========
        ad = ds.get("assessment_date", "")
        if ad:
            ay, am, _ad = _ymd(ad)
            era_a = _era(ay)
            if era_a:
                s(ws, "BI76", era_a)
            if ay: s(ws, "BU75", _era_year(ay))
            if am: s(ws, "CB75", am)
            if _ad: s(ws, "CI75", _ad)

        # ア 病状又は状態像チェックリスト
        _fill_symptoms(ws, ds, sc)

        # イ 具体的記載
        if ds.get("detail_description"):
            s(ws, "BG80", ds["detail_description"])

        # ======== ウ 日常生活状況 ========
        _fill_daily_living(ws, dl, da)

        # ======== エ 就労状況 ========
        _fill_employment(ws, emp)

        # ======== オカキ ========
        if ds.get("physical_findings"):
            s(ws, "BI196", ds["physical_findings"])
        if ds.get("clinical_tests"):
            s(ws, "BI200", ds["clinical_tests"])
        if ds.get("welfare_services"):
            s(ws, "BI205", ds["welfare_services"])

        # ======== ⑪ 日常生活活動能力及び労働能力 ========
        if data.get("daily_ability_and_labor"):
            s(ws, "T210", data["daily_ability_and_labor"])

        # ======== ⑫ 予後 ========
        if data.get("prognosis"):
            s(ws, "T218", data["prognosis"])

        # ======== ⑬ 備考 ========
        if data.get("remarks_13"):
            s(ws, "T225", data["remarks_13"])

        # ======== 署名欄 ========
        if data.get("date"):
            dy, dm, dd = _ymd(data["date"])
            if dy:
                era = _era(dy)
                year_text = _era_year(dy)
                s(ws, "AN233", f"{era} {year_text}" if era else year_text)
            if dm: s(ws, "AW233", dm)
            if dd: s(ws, "BB233", dd)
        if ins.get("name"):
            s(ws, "Z235", ins["name"])
        if ins.get("department"):
            s(ws, "BL235", ins["department"])
        if ins.get("address"):
            s(ws, "Z237", _strip_postal_code(ins["address"]))
        if ins.get("doctor"):
            s(ws, "BL237", ins["doctor"])

        wb.save()
        wb.close()
    finally:
        app.quit()


def _fill_patient_info(ws, patient, reference_date):
    furigana = (
        patient.get("furigana")
        or patient.get("name_kana")
        or patient.get("kana")
        or ""
    )
    if furigana:
        s(ws, "P12", furigana)
    if patient.get("name"):
        s(ws, "P14", patient["name"])

    birthdate = patient.get("birthdate", "")
    by, bm, bd = _ymd(birthdate)
    if birthdate:
        era = _era(by)
        if era:
            s(ws, "BG13", era)
        if by:
            s(ws, "BN12", _era_year(by))
        if bm:
            s(ws, "BT12", bm)
        if bd:
            s(ws, "BZ12", bd)

    age = patient.get("age")
    if age in (None, "") and birthdate:
        age = _calc_age(birthdate, reference_date)
    if age not in (None, ""):
        s(ws, "CI12", str(age))

    sex = str(patient.get("sex", "")).strip()
    if sex in ("男", "男性", "male", "Male", "MALE"):
        s(ws, "CU14", C)
    elif sex in ("女", "女性", "female", "Female", "FEMALE"):
        s(ws, "CZ14", C)

    postal_code = patient.get("postal_code", "") or patient.get("zip_code", "")
    address = patient.get("address", "")
    postcode, address_body = _split_postal_code(postal_code or address)
    if postcode:
        head, tail = postcode.split("-", 1)
        s(ws, "O18", head)
        s(ws, "W18", tail)

    pref, locality = _split_prefecture_address(address_body or address)
    if pref:
        for suffix in ("都", "道", "府", "県"):
            if pref.endswith(suffix):
                s(ws, "AE17", pref[:-1])
                s(ws, "AR17", suffix)
                break
        else:
            s(ws, "AE17", pref)
    if locality:
        city_name, city_suffix, rest = _split_city(locality)
        if city_name:
            s(ws, "AW17", city_name)
        if city_suffix:
            s(ws, "BI17", city_suffix)
        if rest:
            s(ws, "BM17", rest)


def _fill_symptoms(ws, ds, sc):
    """⑩ア 症状チェックリスト — セルに 'レ' を書き込む"""
    prev = ds.get("previous_comparison", "")
    if prev:
        prev_map = {"変化なし": "G81", "改善している": "S81", "悪化している": "AE81", "不明": "AQ81"}
        if prev in prev_map:
            s(ws, prev_map[prev], C)

    dep = sc.get("delusion", [])
    dep_map = {
        "思考・運動制止": "G83", "刺激性、興奮": "W83", "憂うつ気分": "AN83",
        "自殺企図": "G84", "希死念慮": "W84",
        "その他": "G85"
    }
    alt_dep = {"易刺激性、焦燥": "刺激性、興奮", "自殺念慮": "希死念慮"}
    _check_items(ws, dep, dep_map, alt_dep)

    mood = sc.get("mood_state", [])
    mood_map = {
        "行為心迫": "G88", "多弁・多動": "S88", "気分（感情）の異常な高揚・刺激性": "AE88",
        "観念奔逸": "G89", "易怒性・被刺激性亢進": "S89", "誇大妄想": "AQ89",
        "その他": "G90"
    }
    _check_items(ws, mood, mood_map)

    hal = sc.get("hallucination_delusion", [])
    hal_map = {
        "幻覚": "G93", "妄想": "S93", "させられ体験": "AE93", "思考形式の障害": "AQ93",
        "著しい奇異な行為": "G94", "その他": "W94"
    }
    _check_items(ws, hal, hal_map)

    psy = sc.get("psychomotor", [])
    psy_map = {
        "興奮": "G97", "昏迷": "S97", "拒絶・拒食": "AE97", "滅裂思考": "AQ97",
        "衝動行為": "G98", "自傷": "S98", "無動・無反応": "AE98",
        "その他": "G99"
    }
    alt_psy = {"減裂思考": "滅裂思考", "暴発行為": "衝動行為", "無動・無言など": "無動・無反応"}
    _check_items(ws, psy, psy_map, alt_psy)

    res = sc.get("residual_state", [])
    res_map = {
        "自閉": "G102", "感情の平板化": "R102", "意欲の減退": "AE102",
        "その他": "G103"
    }
    _check_items(ws, res, res_map)

    con = sc.get("consciousness", [])
    con_map = {
        "意識混濁": "G106", "(夜間)せん妄": "S106", "もうろう": "AE106", "錯乱": "AQ106",
        "てんかん発作": "G107", "不機嫌症": "S107", "その他": "AE107"
    }
    alt_con = {"せん妄": "(夜間)せん妄", "不機嫌発作": "不機嫌症"}
    _check_items(ws, con, con_map, alt_con)

    epi = sc.get("epilepsy_detail", {})
    if epi:
        stype = epi.get("seizure_type", "")
        if stype:
            type_map = {"A": "W109", "B": "AB109", "C": "AG109", "D": "AL109"}
            if stype in type_map:
                s(ws, type_map[stype], C)
        freq_yr = epi.get("frequency_per_year", "")
        if freq_yr:
            s(ws, "Z110", str(freq_yr))

    intell = sc.get("intellectual", {})
    if intell:
        id_level = intell.get("intellectual_disability", "")
        if id_level:
            s(ws, "G113", C)
            id_map = {"軽度": "R113", "中等度": "Z113", "重度": "AH113", "最重度": "AP113"}
            if id_level in id_map:
                s(ws, id_map[id_level], C)
        dem_level = intell.get("dementia", "")
        if dem_level:
            s(ws, "G114", C)
            dem_map = {"軽度": "R114", "中等度": "Z114", "重度": "AH114", "最重度": "AP114"}
            if dem_level in dem_map:
                s(ws, dem_map[dem_level], C)
        other_items = intell.get("other", [])
        if other_items:
            hcf_map = {
                "失行": "J116", "失認": "T116",
                "記憶障害": "J117", "注意障害": "T117",
                "遂行機能障害": "AD117", "社会的行動障害": "AP117"
            }
            s(ws, "G115", C)
            for item in other_items:
                if item in hcf_map:
                    s(ws, hcf_map[item], C)

    devd = sc.get("developmental_disorder", [])
    devd_map = {
        "相互的な社会関係の質的障害": "G122", "言語コミュニケーションの障害": "AA122",
        "限定した常同的で反復的な関心と行動": "G123", "その他": "AF123"
    }
    _check_items(ws, devd, devd_map)

    pc = sc.get("personality_change", {})
    pc_items = pc.get("items", []) if isinstance(pc, dict) else []
    pc_map = {
        "欠陥状態": "G126", "無関心": "S126", "無為": "AE126",
        "その他": "G127"
    }
    alt_pc = {"攻撃性": "欠陥状態"}
    _check_items(ws, pc_items, pc_map, alt_pc)

    su = sc.get("substance_use", {})
    if isinstance(su, dict):
        substance = su.get("substance", "")
        if substance:
            s(ws, "T129", substance)
        su_items = su.get("items", [])
        su_map = {"乱用": "G130", "依存": "S130"}
        _check_items(ws, su_items, su_map)

    other = sc.get("other", "")
    if other:
        s(ws, "K132", other)


def _check_items(ws, selected, mapping, alternatives=None):
    if not selected:
        return
    for item in selected:
        actual = item
        if alternatives and item in alternatives:
            actual = alternatives[item]
        if actual in mapping:
            s(ws, mapping[actual], C)


def _fill_education(ws, edu):
    level = edu.get("level", "")
    if "普通学級" in level or "普通" in level:
        if "小学校" in level or "小" in level:
            s(ws, "BB54", C)
        if "中学校" in level or "中" in level:
            s(ws, "BB55", C)
        if "高校" in level or "高" in level:
            s(ws, "BB56", C)
    if "特別支援学級" in level:
        s(ws, "BK54", C)
    if "特別支援学校" in level:
        s(ws, "BV54", C)
    if "高校卒業" in level or "高卒" in level:
        s(ws, "BB54", C)
        s(ws, "BB55", C)
        s(ws, "BB56", C)


def _fill_treatment_history(ws, treat_hist):
    rows = [
        ("B60", "U60", "AA60", "AG60", "AM60", "AR61", "AW61", "BB60", "BQ60", "CO60"),
        ("B63", "U63", "AA63", "AG63", "AM63", "AR64", "AW64", "BB63", "BQ63", "CO63"),
        ("B66", "U66", "AA66", "AG66", "AM66", "AR67", "AW67", "BB66", "BQ66", "CO66"),
        ("B69", "U69", "AA69", "AG69", "AM69", "AR70", "AW70", "BB69", "BQ69", "CO69"),
        ("B72", "U72", "AA72", "AG72", "AM72", "AR73", "AW73", "BB72", "BQ72", "CO72"),
    ]
    for i, tr in enumerate(treat_hist[:5]):
        inst, yf, mf, yt, mt, inp, outp, dis, ther, outc = rows[i]
        if tr.get("institution"):
            s(ws, inst, tr["institution"])
        pf = tr.get("period_from", "")
        pt_val = tr.get("period_to", "")
        fy, fm, _ = _ymd(pf)
        ty, tm, _ = _ymd(pt_val)
        if fy: s(ws, yf, _era_year(fy))
        if fm: s(ws, mf, int(fm))
        if ty: s(ws, yt, _era_year(ty))
        if tm: s(ws, mt, int(tm))
        io = tr.get("inpatient_outpatient", "")
        if "入院" in io: s(ws, inp, C)
        if "外来" in io: s(ws, outp, C)
        if tr.get("disease_name"): s(ws, dis, tr["disease_name"])
        if tr.get("treatment"): s(ws, ther, tr["treatment"])
        if tr.get("outcome"): s(ws, outc, tr["outcome"])


def _fill_daily_living(ws, dl, da):
    sit = dl.get("living_situation", "")
    sit_map = {"入院": "K140", "入所": "Q140", "在宅": "W140", "その他": "AC140"}
    if sit in sit_map:
        s(ws, sit_map[sit], C)
    if dl.get("living_detail") and sit == "その他":
        s(ws, "AE140", f"その他（{dl['living_detail']}）")
    elif dl.get("living_detail"):
        s(ws, "AJ140", dl["living_detail"])

    if dl.get("facility_name"):
        s(ws, "R141", dl["facility_name"])

    cohab = dl.get("cohabitant", "")
    if cohab == "有": s(ws, "W142", C)
    elif cohab == "無": s(ws, "AC142", C)

    if dl.get("social_situation"):
        s(ws, "K146", dl["social_situation"])

    assessment_rows = {
        "eating": 154,
        "hygiene": 160,
        "money": 166,
        "medication": 172,
        "communication": 177,
        "safety": 184,
        "social": 190,
    }
    level_cols = {1: "H", 2: "N", 3: "AA", 4: "AQ"}
    for key, row in assessment_rows.items():
        level = da.get(key)
        if level and level in level_cols:
            s(ws, f"{level_cols[level]}{row}", C)

    degree = dl.get("daily_living_level")
    if degree:
        degree_map = {1: "BJ142", 2: "BJ145", 3: "BJ151", 4: "BJ157", 5: "BJ163"}
        if degree in degree_map:
            s(ws, degree_map[degree], C)


def _fill_employment(ws, emp):
    if emp.get("workplace"):
        workplace = emp["workplace"]
        if "一般企業" in workplace:
            s(ws, "O196", C)
        elif "就労支援" in workplace:
            s(ws, "X196", C)
        elif "その他" in workplace:
            s(ws, "AH196", C)
            s(ws, "AO196", workplace.replace("その他", "").replace("（", "").replace("）", "").strip())
        else:
            s(ws, "AH196", C)
            s(ws, "AO196", workplace.strip())

    if emp.get("employment_type"):
        etype = emp["employment_type"]
        if "障害者雇用" in etype: s(ws, "O197", C)
        elif "一般雇用" in etype: s(ws, "X197", C)
        elif "自営" in etype: s(ws, "AH197", C)
        elif etype:
            s(ws, "AM197", C)
            s(ws, "AT197", etype.strip())

    if emp.get("tenure"):
        s(ws, "P198", emp["tenure"])
    if emp.get("work_frequency"):
        s(ws, "AX198", emp["work_frequency"])
    if emp.get("monthly_income"):
        s(ws, "R199", emp["monthly_income"])
    if emp.get("work_content"):
        s(ws, "J201", emp["work_content"])
    if emp.get("support_at_work"):
        s(ws, "J204", emp["support_at_work"])


# ============================================================
# ユーティリティ
# ============================================================
def _add_circle_around_cell(ws, cell_ref):
    """セルの文字を囲む楕円図形を追加する（xlwings経由）"""
    cell = ws.range(cell_ref)
    # セルの位置・サイズを取得
    left = cell.left
    top = cell.top
    width = cell.width
    height = cell.height
    # 楕円を追加（セルに合わせたサイズ、塗りなし・線のみ）
    oval = ws.api.Shapes.AddShape(
        9,  # msoShapeOval
        left + 1, top + 1,
        width - 2, height - 2
    )
    oval.Fill.Visible = False  # 塗りなし
    oval.Line.Weight = 1.0
    oval.Line.ForeColor.RGB = 0  # 黒


def _era_year(western_year_str):
    if not western_year_str:
        return ""
    try:
        y = int(western_year_str)
        if y >= 2019: return str(y - 2018)
        elif y >= 1989: return str(y - 1988)
        elif y >= 1926: return str(y - 1925)
        return str(y)
    except:
        return western_year_str

def _split_postal_code(value):
    if not value:
        return ("", "")
    text = str(value).strip().replace("〒", "")
    if len(text) >= 8 and text[3] == "-" and text[:3].isdigit() and text[4:8].isdigit():
        return (text[:8], text[8:].strip())
    return ("", str(value).strip())

def _strip_postal_code(value):
    postcode, body = _split_postal_code(value)
    return body or str(value).strip()

def _split_city(text):
    if not text:
        return ("", "", "")
    t = str(text).strip()
    for suffix in ("市", "区", "郡"):
        idx = t.find(suffix)
        if idx >= 1:
            return (t[:idx], suffix, t[idx + 1:].strip())
    return ("", "", t)

def _split_prefecture_address(address):
    if not address:
        return ("", "")
    text = str(address).strip()
    for suffix in ("都", "道", "府", "県"):
        idx = text.find(suffix)
        if idx >= 1:
            return (text[:idx + 1], text[idx + 1:].strip())
    return ("", text)

def _calc_age(birthdate, reference_date):
    by, bm, bd = _ymd(birthdate)
    ry, rm, rd = _ymd(reference_date)
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
    for key in ("name", "department", "address", "doctor", "phone",
                "code", "designated_number", "doctor_years"):
        if clinic.get(key):
            merged[key] = clinic[key]
    return merged

def _ymd(d):
    if not d:
        return ("", "", "")
    parts = str(d).split("-")
    if len(parts) >= 3:
        return (parts[0], str(int(parts[1])), str(int(parts[2])))
    if len(parts) == 2:
        return (parts[0], str(int(parts[1])), "")
    return (str(d), "", "")

def _era(year_str):
    if not year_str:
        return ""
    try:
        y = int(year_str)
        if y >= 2019: return "令和"
        elif y >= 1989: return "平成"
        else: return "昭和"
    except:
        return ""


if __name__ == "__main__":
    main()
