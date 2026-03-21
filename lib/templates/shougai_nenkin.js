/**
 * 障害年金診断書（精神の障害用）様式第120号の4 — 公式様式完全再現
 * Usage: node shougai_nenkin.js <input.json> <output.docx>
 */
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageBreak
} = require("docx");

const inputPath = process.argv[2];
const outputPath = process.argv[3];
if (!inputPath || !outputPath) { console.error("Usage: node shougai_nenkin.js <input.json> <output.docx>"); process.exit(1); }
const data = JSON.parse(fs.readFileSync(inputPath, "utf-8"));

// ========== 定数 ==========
const FONT = "Yu Gothic";
const F6 = 12, F7 = 14, F8 = 16, F9 = 18, F10 = 20, F12 = 24, F14 = 28;
const thin = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
const bdr = { top: thin, bottom: thin, left: thin, right: thin };

function t(text, o = {}) {
  return new TextRun({ text: text || "", font: o.font || FONT, size: o.s || F8, bold: !!o.b, color: o.c });
}
function p(runs, o = {}) {
  const ch = Array.isArray(runs) ? runs : [t(runs, o)];
  return new Paragraph({
    alignment: o.a || AlignmentType.LEFT,
    spacing: { before: o.bf || 0, after: o.af || 0, line: o.l || 220 },
    indent: o.indent, children: ch,
  });
}
function mc(content, w, o = {}) {
  const ch = Array.isArray(content) ? content
    : typeof content === "string" ? [p(content, { s: o.fs || F8 })] : [content];
  const r = {
    borders: o.bdr || bdr, width: { size: w, type: WidthType.DXA },
    margins: o.m || { top: 10, bottom: 10, left: 40, right: 40 },
    verticalAlign: o.v || VerticalAlign.TOP, children: ch,
  };
  if (o.bg) r.shading = { fill: o.bg, type: ShadingType.CLEAR };
  if (o.cs) r.columnSpan = o.cs;
  if (o.rs) r.rowSpan = o.rs;
  return new TableCell(r);
}
function hc(text, w, o = {}) {
  return mc([p(text, { s: o.fs || F7, a: AlignmentType.CENTER, l: 210 })], w, { ...o, bg: "F0F0F0" });
}
function ml(text, w, o = {}) {
  return mc((text || "").split("\n").map(l => p(l, { s: o.fs || F8, l: 230 })), w, o);
}

function ci(arr, text) { return (arr || []).includes(text) ? `○${text}` : `　${text}`; }
function cv(val, text) { return val === text ? `○${text}` : `　${text}`; }

function fd(d) {
  if (!d) return "";
  const ps = d.split("-");
  if (ps.length >= 3) return `${ps[0]}年${parseInt(ps[1])}月${parseInt(ps[2])}日`;
  if (ps.length === 2) return `${ps[0]}年${parseInt(ps[1])}月`;
  return d;
}
function fdb(d) { return d ? fd(d) : "　　年　　月　　日"; }

// ========== メイン ==========
async function generate() {
  const pt = data.patient || {};
  const diag = data.diagnosis || {};
  const fv = data.first_visit || {};
  const rem = data.remission || {};
  const hist = data.history || {};
  const init = data.initial_findings || {};
  const dev = data.development_history || {};
  const ds = data.disability_state || {};
  const sc = ds.symptom_checklist || {};
  const dl = ds.daily_living || {};
  const da = dl.daily_assessment || {};
  const emp = ds.employment || {};
  const ins = data.institution || {};

  const TW = 10206;
  const HL = 2200; // ヘッダー列幅

  // ===== ⑩ア: 病状又は状態像チェックリスト =====
  function buildSymptoms() {
    const rows = [];
    const line = (label, arr, items) => {
      rows.push(p([t(label, { s: F7, b: true })], { af: 3 }));
      // アイテムを1行にまとめる
      let itemTexts = items.map(item => ci(arr, item));
      rows.push(p([t("　" + itemTexts.join("　"), { s: F7 })], { af: 8 }));
    };

    // 前回比較
    if (ds.previous_comparison) {
      rows.push(p([t("前回の診断書の記載時との比較（前回の診断書を作成している場合は記入してください。）", { s: F6 })], { af: 3 }));
      rows.push(p([t(`　${cv(ds.previous_comparison,"変化なし")}　${cv(ds.previous_comparison,"改善している")}　${cv(ds.previous_comparison,"悪化している")}　${cv(ds.previous_comparison,"不明")}`, { s: F7 })], { af: 10 }));
    }

    // I 抑うつ状態
    rows.push(p([t("Ⅰ　抑うつ状態", { s: F7, b: true })], { af: 3 }));
    const dep = sc.delusion || [];
    rows.push(p([t(`　${ci(dep,"思考・運動制止")}　${ci(dep,"易刺激性、焦燥")}　${ci(dep,"憂うつ気分")}`, { s: F7 })], { af: 2 }));
    rows.push(p([t(`　${ci(dep,"自殺念慮")}　${ci(dep,"自殺企図")}`, { s: F7 })], { af: 2 }));
    rows.push(p([t(`　${ci(dep,"その他")}（${(dep.includes("その他") && sc.delusion_other) || "　　　　　"}）`, { s: F7 })], { af: 8 }));

    // II そう状態
    rows.push(p([t("Ⅱ　そう状態", { s: F7, b: true })], { af: 3 }));
    const mood = sc.mood_state || [];
    rows.push(p([t(`　${ci(mood,"行為心迫")}　${ci(mood,"多弁・多動")}　${ci(mood,"気分（感情）の異常な高揚・刺激性")}`, { s: F7 })], { af: 2 }));
    rows.push(p([t(`　${ci(mood,"観念奔逸")}　${ci(mood,"易怒性・被刺激性亢進")}　${ci(mood,"誇大妄想")}`, { s: F7 })], { af: 8 }));

    // III 幻覚妄想状態
    rows.push(p([t("Ⅲ　幻覚妄想状態", { s: F7, b: true })], { af: 3 }));
    const hal = sc.hallucination_delusion || [];
    rows.push(p([t(`　${ci(hal,"幻覚")}　${ci(hal,"妄想")}　${ci(hal,"させられ体験")}　${ci(hal,"思考形式の障害")}`, { s: F7 })], { af: 2 }));
    rows.push(p([t(`　${ci(hal,"著しい奇異な行為")}　${ci(hal,"その他")}（${(hal.includes("その他") && sc.hallucination_other) || "　　　　"}）`, { s: F7 })], { af: 8 }));

    // IV 精神運動興奮状態及び停滞の状態
    rows.push(p([t("Ⅳ　精神運動興奮状態及び停滞の状態", { s: F7, b: true })], { af: 3 }));
    const psy = sc.psychomotor || [];
    rows.push(p([t(`　${ci(psy,"興奮")}　${ci(psy,"昏迷")}　${ci(psy,"拒絶・拒食")}　${ci(psy,"減裂思考")}`, { s: F7 })], { af: 2 }));
    rows.push(p([t(`　${ci(psy,"暴発行為")}　${ci(psy,"自傷")}　${ci(psy,"無動・無言など")}　${ci(psy,"その他")}（${(psy.includes("その他") && sc.psychomotor_other) || "　　　"}）`, { s: F7 })], { af: 8 }));

    // V 統合失調症等残遺状態
    rows.push(p([t("Ⅴ　統合失調症等残遺状態", { s: F7, b: true })], { af: 3 }));
    const res = sc.residual_state || [];
    rows.push(p([t(`　${ci(res,"自閉")}　${ci(res,"感情の平板化")}　${ci(res,"意欲の減退")}　${ci(res,"その他")}（${(res.includes("その他") && sc.residual_other) || "　　　"}）`, { s: F7 })], { af: 8 }));

    // VI 意識障害等（てんかんを含む。）
    rows.push(p([t("Ⅵ　意識障害等（てんかんを含む。）", { s: F7, b: true })], { af: 3 }));
    const con = sc.consciousness || [];
    rows.push(p([t(`　${ci(con,"せん妄")}　${ci(con,"もうろう")}　${ci(con,"錯乱")}`, { s: F7 })], { af: 2 }));
    rows.push(p([t(`　${ci(con,"てんかん発作")}　${ci(con,"不機嫌発作")}　${ci(con,"その他")}（${(con.includes("その他") && sc.consciousness_other) || "　　　"}）`, { s: F7 })], { af: 5 }));

    // てんかん詳細
    rows.push(p([t("　・てんかん及び意識障害　※発作のタイプは以下の注を参照", { s: F6 })], { af: 3 }));
    const epi = sc.epilepsy_detail || {};
    rows.push(p([t(`　　てんかん発作のタイプ（${epi.seizure_type || "Ａ・Ｂ・Ｃ・Ｄ"}）`, { s: F7 })], { af: 2 }));
    rows.push(p([t(`　　てんかん発作の頻度（年間${epi.frequency_per_year || "　　"}回、月平均${epi.frequency_per_month || "　　"}回）`, { s: F7 })], { af: 2 }));
    rows.push(p([t(`　　最終発作　${epi.last_seizure ? fd(epi.last_seizure) : "　　年　　月　　日"}`, { s: F7 })], { af: 8 }));

    // VII 知能障害等
    rows.push(p([t("Ⅶ　知能障害等", { s: F7, b: true })], { af: 3 }));
    const intl = sc.intellectual || {};
    rows.push(p([t(`　１　知的障害　　ア軽度　イ中等度　ウ重度　エ最重度`, { s: F7 })], { af: 2 }));
    rows.push(p([t(`　２　認知症　　　ア軽度　イ中等度　ウ重度　エ最重度`, { s: F7 })], { af: 2 }));
    rows.push(p([t("　３　高次脳機能障害", { s: F7 })], { af: 2 }));
    rows.push(p([t("　　　ア失行　イ失認　ウ記憶障害　エ注意障害　オ遂行機能障害　カ社会的行動障害", { s: F7 })], { af: 2 }));
    rows.push(p([t(`　４　学習障害　ア読み　イ書き　ウ計算　エその他（${intl.learning_other || "　　　"}）`, { s: F7 })], { af: 2 }));
    rows.push(p([t(`　５　その他（${(intl.other || []).join("、") || "　　　　　　"}）`, { s: F7 })], { af: 8 }));

    // VIII 発達障害関連症状
    rows.push(p([t("Ⅷ　発達障害関連症状", { s: F7, b: true })], { af: 3 }));
    const devd = sc.developmental_disorder || [];
    rows.push(p([t(`　${ci(devd,"相互的な社会関係の質的障害")}　${ci(devd,"言語コミュニケーションの障害")}`, { s: F7 })], { af: 2 }));
    rows.push(p([t(`　${ci(devd,"限定した常同的で反復的な関心と行動")}　${ci(devd,"その他")}（${(devd.includes("その他") && sc.developmental_other) || "　　　"}）`, { s: F7 })], { af: 8 }));

    // IX 人格変化
    rows.push(p([t("Ⅸ　人格変化", { s: F7, b: true })], { af: 3 }));
    const pc = (sc.personality_change || {}).items || [];
    rows.push(p([t(`　${ci(pc,"攻撃性")}　${ci(pc,"無関心")}　${ci(pc,"無為")}`, { s: F7 })], { af: 2 }));
    rows.push(p([t(`　${ci(pc,"その他")}（${(pc.includes("その他") && sc.personality_other) || "　　　　　"}）`, { s: F7 })], { af: 8 }));

    // X 乱用、依存等
    rows.push(p([t("Ⅹ　乱用、依存等（薬物等名；", { s: F7, b: true }), t(`${(sc.substance_use || {}).substance || "　　　　"}）`, { s: F7 })], { af: 3 }));
    const su = (sc.substance_use || {}).items || [];
    rows.push(p([t(`　${ci(su,"乱用")}　${ci(su,"依存")}`, { s: F7 })], { af: 8 }));

    // XI その他
    rows.push(p([t("Ⅺ　その他〔", { s: F7, b: true }), t(`${sc.other || "　　　　　　　　　　　　　　　　　　"}`, { s: F7 }), t("〕", { s: F7, b: true })], { af: 5 }));

    rows.push(p([t("本人の障害の程度及び状態に無関係な欄には記入する必要はありません。（無関係な欄は、斜線により抹消してください。）", { s: F6, c: "666666" })], { af: 5 }));

    return rows;
  }

  // ===== 治療歴テーブル行 =====
  const th = dev.treatment_history || [];
  const emptyTreat = { institution: "", period_from: "", period_to: "", inpatient_outpatient: "入院・外来", disease_name: "", treatment: "", outcome: "" };
  const treatData = th.length > 0 ? th : [emptyTreat, emptyTreat, emptyTreat, emptyTreat];
  const BW = TW - HL; // body width
  const treatRows = treatData.map(tr => new TableRow({ children: [
    mc(tr.institution || "", 1700, { fs: F7 }),
    mc([p([t(`${tr.period_from || "　年　月"}～${tr.period_to || "　年　月"}`, { s: F7 })])], 1400),
    mc(tr.inpatient_outpatient || "入院・外来", 800, { fs: F7 }),
    mc(tr.disease_name || "", 1400, { fs: F7 }),
    mc(tr.treatment || "", 1600, { fs: F7 }),
    mc(tr.outcome ? tr.outcome : "軽快・悪化・不変", BW - 1700 - 1400 - 800 - 1400 - 1600, { fs: F7 }),
  ]}));

  // ===== 日常生活能力の判定行 =====
  function dailyRow(label, desc, key) {
    const val = da[key] || 0;
    return new TableRow({ children: [
      mc([
        p([t(label, { s: F7, b: true })], { af: 3 }),
        p([t(desc, { s: F6 })]),
      ], 2400),
      mc([p([
        t(val === 1 ? "☑" : "□", { s: F8 }), t("できる　", { s: F7 }),
        t(val === 2 ? "☑" : "□", { s: F8 }), t("おおむねできるが\n援助が必要な場面もある　", { s: F6 }),
        t(val === 3 ? "☑" : "□", { s: F8 }), t("自発的にはできないが\n援助があればできる　", { s: F6 }),
        t(val === 4 ? "☑" : "□", { s: F8 }), t("できない", { s: F6 }),
      ], { l: 200 })], TW - 2400),
    ]});
  }

  // ===== 日常生活能力の程度テキスト =====
  const dlLabels = {
    1: "（1）精神障害（病的体験・残遺症状・認知障害・性格変化等）を認めるが、社会生活は普通にできる。",
    2: "（2）精神障害を認め、家庭内での日常生活は普通にできるが、社会生活には、援助が必要である。",
    3: "（3）精神障害を認め、家庭内での単純な日常生活はできるが、時に応じて援助が必要である。",
    4: "（4）精神障害を認め、日常生活における身のまわりのことも、多くの援助が必要である。",
    5: "（5）精神障害を認め、身のまわりのこともほとんどできないため、常時の援助が必要である。",
  };

  // ========== ドキュメント構築 ==========
  const doc = new Document({
    styles: { default: { document: { run: { font: FONT, size: F8 } } } },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 450, right: 650, bottom: 350, left: 650 },
        },
      },
      children: [
        // ===== ページ1 =====
        // タイトル
        p([t("国民年金", { s: F9 }), t("　　　　　　　　　　", { s: F7 }), ], { a: AlignmentType.CENTER }),
        p([t("厚生年金保険", { s: F9 }), t("　　　診　断　書", { s: F14, b: true }), t("（精神の障害用）", { s: F9 })], { a: AlignmentType.CENTER, af: 5 }),
        p([t("様式第120号の4", { s: F7 })], { a: AlignmentType.RIGHT, af: 30 }),

        // --- 患者情報 ---
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("氏　名", 1000),
            mc(pt.name || "", 3400, { fs: F10 }),
            hc("生年月日", 1000),
            mc([p([t("昭和・平成・令和", { s: F6 }), t(pt.birthdate ? `　${fd(pt.birthdate)}生` : "　　年　月　日生", { s: F7 }), t(pt.age != null ? `（${pt.age}歳）` : "（　歳）", { s: F7 })])], 2406),
            hc("性別", 600),
            mc([p([t(pt.sex ? (pt.sex === "男" ? "○男・女" : "男・○女") : "男・女", { s: F8 })])], TW - 1000 - 3400 - 1000 - 2406 - 600),
          ]}),
          new TableRow({ children: [
            hc("住　所", 1000),
            mc([p([t("〒", { s: F7 }), t(pt.postal_code || "　　　-　　　　", { s: F7 })]), p(pt.address || "", { s: F8 })], TW - 1000, { cs: 5 }),
          ]}),
        ]}),

        p("", { af: 15 }),

        // --- ①〜⑥ ---
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          // ① 障害の原因となった傷病名
          new TableRow({ children: [
            hc("①\n障害の原因と\nなった傷病名", HL),
            mc([
              p([t(diag.disease_name || "", { s: F9 })], { af: 5 }),
              p([t("ＩＣＤ－１０コード（", { s: F7 }), t(diag.icd_code || "　　　　", { s: F8 }), t("）", { s: F7 })]),
            ], TW - HL),
          ]}),

          // ② 傷病の発生年月日
          new TableRow({ children: [
            hc("②　傷病の発生\n年月日", HL),
            mc([p([
              t("平成・令和", { s: F6 }),
              t(diag.onset_date ? `　${fd(diag.onset_date)}` : "　　年　　月　　日", { s: F8 }),
              t("　　", { s: F7 }),
              t("診療録で確定", { s: F7 }),
              t("　", { s: F7 }),
              t(diag.onset_source === "本人の申立て" ? "○本人の申立て" : "　本人の申立て", { s: F7 }),
            ]), p([
              t(diag.occupation_at_onset ? `本人の発病時の職業：${diag.occupation_at_onset}` : "本人の発病時の職業：", { s: F7 }),
            ])], TW - HL),
          ]}),

          // ③ 初めて医師の診療を受けた日
          new TableRow({ children: [
            hc("③　①のため初めて\n医師の診療を\n受けた日", HL),
            mc([p([
              t("昭和・平成・令和", { s: F6 }),
              t(fv.date ? `　${fd(fv.date)}` : "　　年　　月　　日", { s: F8 }),
              t("　　", { s: F7 }),
              t("診療録で確定", { s: F7 }),
              t("　", { s: F7 }),
              t(fv.source === "本人の申立て" ? "○本人の申立て" : "　本人の申立て", { s: F7 }),
            ])], TW - HL),
          ]}),

          // ⑥ 傷病が治ったかどうか
          new TableRow({ children: [
            hc("⑥　傷病が治った\n（症状が固定した\n状態を含む。）\nかどうか", HL),
            mc([
              p([
                t("平成・令和", { s: F6 }),
                t(rem.remission_date ? `　${fd(rem.remission_date)}` : "　　年　　月　　日", { s: F8 }),
                t("　", { s: F7 }),
                t(rem.confirmed_or_estimated || "確認・推定", { s: F7 }),
              ]),
              p([
                t("症状のよくなる見込み・・・", { s: F7 }),
                t(`${cv(rem.prognosis_outlook,"有")}　・　${cv(rem.prognosis_outlook,"無")}　・　${cv(rem.prognosis_outlook,"不明")}`, { s: F7 }),
              ]),
            ], TW - HL),
          ]}),

          // ④ 既存障害
          new TableRow({ children: [
            hc("④既存障害", HL),
            mc(data.existing_disability || "", TW - HL),
          ]}),

          // ⑤ 既往症
          new TableRow({ children: [
            hc("⑤既往症", HL),
            mc(data.prior_illness || "", TW - HL),
          ]}),
        ]}),

        p("", { af: 10 }),

        // --- ⑦ 病歴 ---
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("⑦", HL, { rs: 2 }),
            mc([
              p([t("陳述者の氏名", { s: F7 }), t(`　${hist.informant_name || ""}`, { s: F8 })], { af: 5 }),
            ], 2600),
            mc([
              p([t("請求人との続柄", { s: F7 }), t(`　${hist.informant_relationship || ""}`, { s: F8 })]),
            ], 2200),
            mc([
              p([t("聴取年月日", { s: F7 }), t(`　${hist.interview_date ? fd(hist.interview_date) : "　年　月　日"}`, { s: F8 })]),
            ], TW - HL - 2600 - 2200),
          ]}),
          new TableRow({ children: [
            mc([
              p([t("発病から現在までの病歴及び治療の経過、内容、就学・就労状況等、期間、その他参考となる事項", { s: F6, c: "666666" })], { af: 10 }),
              ...(hist.narrative || "").split("\n").map(l => p(l, { s: F7, l: 230 })),
            ], TW - HL, { cs: 3 }),
          ]}),
        ]}),

        p("", { af: 10 }),

        // --- ⑧ 初診時所見 ---
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("⑧\n診断書作成医療機関\nにおける初診時所見", HL),
            mc([
              p([t("初診年月日", { s: F7 }), t("　昭和・平成・令和", { s: F6 }), t(init.first_visit_date ? `　${fd(init.first_visit_date)}` : "　　年　月　日", { s: F7 })], { af: 10 }),
              ...(init.findings || "").split("\n").map(l => p(l, { s: F7, l: 230 })),
            ], TW - HL),
          ]}),
        ]}),

        p("", { af: 10 }),

        // --- ⑨ 発育・養育歴等 ---
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("⑨", HL, { rs: 3 }),
            hc("ア 発育・養育歴", 1500),
            ml(dev.development || "", TW - HL - 1500),
          ]}),
          new TableRow({ children: [
            hc("イ 教育歴", 1500),
            mc([p([
              t((dev.education || {}).level || "", { s: F7 }),
              t((dev.education || {}).detail ? `　${dev.education.detail}` : "", { s: F7 }),
            ])], TW - HL - 1500),
          ]}),
          new TableRow({ children: [
            hc("ウ 職歴", 1500),
            ml(dev.work_history || "", TW - HL - 1500),
          ]}),
        ]}),

        p("", { af: 5 }),

        // エ 治療歴
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("エ 治療歴", HL),
            hc("医療機関名", 1700),
            hc("治療期間", 1400),
            hc("入院・外来", 800),
            hc("病　名", 1400),
            hc("主な療法", 1600),
            hc("転帰\n(軽快・悪化・不変)", BW - 1700 - 1400 - 800 - 1400 - 1600),
          ]}),
          ...treatRows,
        ]}),

        // ページ区切り
        new Paragraph({ children: [new PageBreak()] }),

        // ===== ページ2 =====

        // --- ⑩ 障害の状態 ---
        p([t("⑩　　障　害　の　状　態", { s: F10, b: true }),
          t(`　（　${ds.assessment_date ? fd(ds.assessment_date).replace("年","・").replace("月","・").replace("日","") : "平成・令和　　年　　月　　日"}　現症　）`, { s: F8 })], { af: 10 }),

        // ⑩ア 左右2列: ア=左、イ=右
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            mc([
              p([t("ア　現在の病状又は状態像", { s: F8, b: true })], { af: 5 }),
              p([t("（該当のローマ数字、英数字を○で囲んでください。）", { s: F6, c: "666666" })], { af: 10 }),
              ...buildSymptoms(),
            ], Math.floor(TW * 0.55)),
            mc([
              p([t("イ　左記の状態について、その程度・症状・処方薬等を具体的に記載してください。", { s: F7, b: true })], { af: 10 }),
              ...(ds.detail_description || "").split("\n").map(l => p(l, { s: F7, l: 230 })),
            ], TW - Math.floor(TW * 0.55)),
          ]}),
        ]}),

        // ページ区切り
        new Paragraph({ children: [new PageBreak()] }),

        // ===== ページ3 =====

        // --- ウ 日常生活状況 ---
        p([t("ウ　日常生活状況", { s: F9, b: true })], { af: 10 }),
        p([t("１　家庭及び社会生活についての具体的な状況", { s: F8 })], { af: 5 }),

        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("（ア）現在の生活環境\n（該当するものの一つを\n○で囲んでください。）", 2400),
            mc([
              p([t(`${cv(dl.living_situation,"入院")}　、　${cv(dl.living_situation,"入所")}　、　${cv(dl.living_situation,"在宅")}　、　${cv(dl.living_situation,"その他")}（${dl.living_detail || "　　　"}）`, { s: F7 })]),
              p([t("（施設名", { s: F7 }), t(`　${dl.facility_name || "　　　　　　　　　　　　"}`, { s: F7 }), t("）", { s: F7 })]),
              p([t(`同居者の有無　（　${cv(dl.cohabitant,"有")}　・　${cv(dl.cohabitant,"無")}　）`, { s: F7 })]),
            ], TW - 2400),
          ]}),
          new TableRow({ children: [
            hc("（イ）全般的状況\n（家族及び家族以外の者\nとの対人関係について\n具体的に記入してください。）", 2400),
            ml(dl.social_situation || "", TW - 2400, { fs: F7 }),
          ]}),
        ]}),

        p("", { af: 10 }),

        // --- 2 日常生活能力の判定 ---
        p([t("２　日常生活能力の判定（該当するものにチェックしてください。）", { s: F8, b: true })], { af: 3 }),
        p([t("（判断にあたっては、単身で生活するとしたら可能かどうかで判断してください。）", { s: F7 })], { af: 10 }),

        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          dailyRow("（1）適切な食事", "—配膳などの準備も含めて適当量をバランスよくとることがなど。", "eating"),
          dailyRow("（2）身辺の清潔保持", "—洗面、洗髪、入浴等の身体の衛生保持や着替え等ができるか。また、自室の清掃や片付けができるなど。", "hygiene"),
          dailyRow("（3）金銭管理と買い物", "—金銭を独力で適切に管理し、やりくりがはかれるか。また一人で買い物が可能であり、計画的な買い物ができるなど。", "money"),
          dailyRow("（4）通院と服薬（要・不要）", "—規則的に通院や服薬を行い、病状等を主治医に伝えることができるなど。", "medication"),
          dailyRow("（5）他人との意思伝達及び対人関係", "—他人の話を聞く、自分の意思を相手に伝える、集団的行動が行えるなど。", "communication"),
          dailyRow("（6）身辺の安全保持及び危機対応", "—事故等の危険から身を守る能力がある、通常と異なる事態となった時に他人に援助を求めるなどを含めて、適正に対応することができるなど。", "safety"),
          dailyRow("（7）社会性", "—銀行での金銭の出し入れや公共施設等の利用が一人で可能。また、社会生活に必要な手続きが行えるなど。", "social"),
        ]}),

        p("", { af: 10 }),

        // --- 3 日常生活能力の程度 ---
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("３　日常生活能力の程度\n（該当するものの一つを\n○で囲んでください。）", 2400),
            mc([
              p([t("（精神障害）", { s: F7, b: true })], { af: 5 }),
              ...[1,2,3,4,5].map(n =>
                p([t(dl.daily_living_level === n ? `○${dlLabels[n]}` : `　${dlLabels[n]}`, { s: F7 })], { af: 5 })
              ),
            ], TW - 2400),
          ]}),
        ]}),

        p("", { af: 10 }),

        // --- エ 就労状況 ---
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("エ　現症時の就労状況", 2400),
            mc([
              p([t(`○勤務先　、　一般企業　、　就労支援施設　、　その他（${emp.workplace || "　　　　"}）`, { s: F7 })]),
              p([t(`○雇用体系　、　障害者雇用　、　一般雇用　、　自営　、　その他（${emp.employment_type || "　　　"}）`, { s: F7 })]),
              p([t(`○勤続年数（${emp.tenure || "　　"}年　＋月）　○仕事の頻度（週に・月に（${emp.work_frequency || "　　"}）日）`, { s: F7 })]),
              p([t(`○ひと月の給与（${emp.monthly_income || "　　　　"}円程度）`, { s: F7 })]),
              p([t(`○仕事の内容　${emp.work_content || ""}`, { s: F7 })]),
              p([t("○仕事場での援助の状況や意思疎通の状況", { s: F7 })]),
              ...(emp.support_at_work || "").split("\n").map(l => p(l, { s: F7, l: 230 })),
            ], TW - 2400),
          ]}),
        ]}),

        p("", { af: 8 }),

        // --- オ カ キ ---
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("オ　身体所見\n（神経学的な所見を含む。）", 2400),
            ml(ds.physical_findings || "", TW - 2400, { fs: F7 }),
          ]}),
          new TableRow({ children: [
            hc("カ　臨床検査\n（心理テスト・認知検査、\n知能障害の場合は、知能指数、\n精神年齢を含む。）", 2400),
            ml(ds.clinical_tests || "", TW - 2400, { fs: F7 }),
          ]}),
          new TableRow({ children: [
            hc("キ　福祉サービスの利用状況\n（障害者総合支援法に規定する\n自立訓練、共同生活援助、\n居宅介護、その他事業福祉\nサービス等）", 2400),
            ml(ds.welfare_services || "", TW - 2400, { fs: F7 }),
          ]}),
        ]}),

        p("", { af: 10 }),

        // --- ⑪ ⑫ ⑬ ---
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("⑪\n現症時の日常生活\n活動能力及び労働能力\n（必ず記入してください。）", 2400),
            ml(data.daily_ability_and_labor || "", TW - 2400),
          ]}),
          new TableRow({ children: [
            hc("⑫\n予　　後\n（必ず記入してください。）", 2400),
            ml(data.prognosis || "", TW - 2400),
          ]}),
          new TableRow({ children: [
            hc("⑬\n備　　考", 2400),
            ml(data.remarks_13 || "", TW - 2400),
          ]}),
        ]}),

        p("", { af: 30 }),

        // --- 署名欄 ---
        p([t("上記のとおり、診断します。", { s: F8 })], { af: 15 }),
        p([t(data.date ? fd(data.date) : "　　年　　月　　日", { s: F9 })], { a: AlignmentType.RIGHT, af: 20 }),

        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("病院又は診療所の名称", 2400),
            mc(ins.name || "", 4000),
            hc("診療担当科名", 1600),
            mc(ins.department || "", TW - 2400 - 4000 - 1600),
          ]}),
          new TableRow({ children: [
            hc("所　在　地", 2400),
            mc(ins.address || "", TW - 2400, { cs: 3 }),
          ]}),
          new TableRow({ children: [
            hc("医師氏名", 2400),
            mc([p([t(ins.doctor || "", { s: F10 })])], TW - 2400, { cs: 3 }),
          ]}),
        ]}),
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log(`生成完了: ${outputPath}`);
}

generate().catch(err => { console.error("生成エラー:", err); process.exit(1); });
