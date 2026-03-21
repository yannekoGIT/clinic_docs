/**
 * 自立支援医療診断書（精神通院）— 第12号様式（第８条関係）完全再現
 * Usage: node jiritsu_shien.js <input.json> <output.docx>
 */
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign
} = require("docx");

const inputPath = process.argv[2];
const outputPath = process.argv[3];
if (!inputPath || !outputPath) { console.error("Usage: node jiritsu_shien.js <input.json> <output.docx>"); process.exit(1); }
const data = JSON.parse(fs.readFileSync(inputPath, "utf-8"));

// ========== 定数・ユーティリティ ==========
const FONT = "Yu Gothic";
const F8 = 16, F7 = 14, F9 = 18, F10 = 20, F12 = 24;
const thin = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
const bdr = { top: thin, bottom: thin, left: thin, right: thin };
const noBdr = { style: BorderStyle.NONE, size: 0 };
const noBdrs = { top: noBdr, bottom: noBdr, left: noBdr, right: noBdr };

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
function c(content, w, o = {}) {
  const ch = Array.isArray(content) ? content
    : typeof content === "string" ? [p(content, { s: o.fs || F8 })] : [content];
  const r = {
    borders: o.b || bdr, width: { size: w, type: WidthType.DXA },
    margins: o.m || { top: 15, bottom: 15, left: 40, right: 40 },
    verticalAlign: o.v || VerticalAlign.TOP, children: ch,
  };
  if (o.bg) r.shading = { fill: o.bg, type: ShadingType.CLEAR };
  if (o.cs) r.columnSpan = o.cs;
  if (o.rs) r.rowSpan = o.rs;
  return new TableCell(r);
}
function hc(text, w, o = {}) {
  return c([p(text, { s: o.fs || F8, a: AlignmentType.CENTER, l: 220 })], w, { ...o, bg: "F0F0F0" });
}

// ○ for selected items
function ci(arr, text) { return (arr || []).includes(text) ? `○${text}` : `　${text}`; }
// ○ for single value
function cv(val, text) { return val === text ? `○${text}` : `　${text}`; }

function formatDate(d) {
  if (!d) return "";
  const ps = d.split("-");
  if (ps.length >= 3) return `${ps[0]}年${parseInt(ps[1])}月${parseInt(ps[2])}日`;
  if (ps.length === 2) return `${ps[0]}年${parseInt(ps[1])}月`;
  return d;
}

// ========== メイン ==========
async function generate() {
  const pt = data.patient || {};
  const cd = data.current_condition || {};
  const tr = data.treatment || {};
  const ins = data.institution || {};
  const epi = cd.epilepsy || {};
  const hc_data = data.heavy_continuous || {};
  const ws = data.welfare_services || {};
  const wsItems = ws.items || [];

  const TW = 10206;

  // --- 薬剤テキスト ---
  const medsLines = (tr.medications || []).map(m =>
    [m.name, m.dosage, m.frequency].filter(Boolean).join("　")
  );

  // --- セクション3: 全12カテゴリのチェック項目テキスト生成 ---
  function sec3() {
    const rows = [];
    const dep = cd.depressive_state || [];
    const man = cd.manic_state || [];
    const hal = cd.hallucination_delusion || [];
    const psy = cd.psychomotor_excitement || cd.cognitive_decline || [];
    const res = cd.residual_state || cd.personality_behavior || [];
    const emo = cd.emotion_behavior || [];
    const anx = cd.anxiety_neurosis || [];
    const sub = cd.substance_use || [];
    const cog = cd.cognitive_learning || [];
    const dev = cd.developmental || [];

    // (1) 抑うつ状態
    rows.push(p([t("（1）　抑うつ状態", { s: F8, b: true })]));
    rows.push(p([t(`　　${ci(dep,"思考・運動抑制")}　${ci(dep,"易刺激性・興奮")}　${ci(dep,"憂うつ気分")}　${ci(dep,"その他")}`, { s: F7 }),
      t(dep.includes("その他") && cd.depressive_other ? `（${cd.depressive_other}）` : "（　　　　　　　　）", { s: F7 })], { af: 10 }));

    // (2) 躁状態
    rows.push(p([t("（2）　躁状態", { s: F8, b: true })]));
    rows.push(p([t(`　　${ci(man,"行為心迫")}　${ci(man,"多弁")}　${ci(man,"感情高揚・易刺激性")}　${ci(man,"その他")}`, { s: F7 }),
      t(man.includes("その他") && cd.manic_other ? `（${cd.manic_other}）` : "（　　　　　　　　）", { s: F7 })], { af: 10 }));

    // (3) 幻覚妄想状態
    rows.push(p([t("（3）　幻覚妄想状態", { s: F8, b: true })]));
    rows.push(p([t(`　　${ci(hal,"幻覚")}　${ci(hal,"妄想")}　${ci(hal,"その他")}`, { s: F7 }),
      t(hal.includes("その他") && cd.hallucination_other ? `（${cd.hallucination_other}）` : "（　　　　　　　　）", { s: F7 })], { af: 10 }));

    // (4) 精神運動興奮及び昏迷の状態
    rows.push(p([t("（4）　精神運動興奮及び昏迷の状態", { s: F8, b: true })]));
    rows.push(p([t(`　　${ci(psy,"興奮")}　${ci(psy,"昏迷")}　${ci(psy,"拒絶")}　${ci(psy,"その他")}`, { s: F7 }),
      t(psy.includes("その他") && cd.psychomotor_other ? `（${cd.psychomotor_other}）` : "（　　　　　　　　）", { s: F7 })], { af: 10 }));

    // (5) 統合失調症等残遺状態
    rows.push(p([t("（5）　統合失調症等残遺状態", { s: F8, b: true })]));
    rows.push(p([t(`　　${ci(res,"自閉")}　${ci(res,"感情平板化")}　${ci(res,"意欲の減退")}　${ci(res,"その他")}`, { s: F7 }),
      t(res.includes("その他") && cd.residual_other ? `（${cd.residual_other}）` : "（　　　　　　　　）", { s: F7 })], { af: 10 }));

    // (6) 情動及び行動の障害
    rows.push(p([t("（6）　情動及び行動の障害", { s: F8, b: true })]));
    rows.push(p([t(`　　${ci(emo,"爆発性")}　${ci(emo,"易怒性")}　${ci(emo,"気分変動")}　${ci(emo,"暴力・衝動行為")}　${ci(emo,"常同行為")}　${ci(emo,"多動")}`, { s: F7 })]));
    rows.push(p([t(`　　${ci(emo,"食行動の異常")}　${ci(emo,"性行動の異常")}　${ci(emo,"チック・汚言")}　${ci(emo,"その他")}`, { s: F7 }),
      t(emo.includes("その他") && cd.emotion_other ? `（${cd.emotion_other}）` : "（　　　　　　　　）", { s: F7 })], { af: 10 }));

    // (7) 不安及び不穏状態
    rows.push(p([t("（7）　不安及び不穏状態", { s: F8, b: true })]));
    rows.push(p([t(`　　${ci(anx,"強度の不安・恐怖感")}　${ci(anx,"精神運動不穏")}　${ci(anx,"心身衰弱")}　${ci(anx,"強迫体験")}　${ci(anx,"心気症状")}`, { s: F7 })]));
    rows.push(p([t(`　　${ci(anx,"心的外傷に関連する症状")}　${ci(anx,"解離・転換症状")}　${ci(anx,"その他")}`, { s: F7 }),
      t(anx.includes("その他") && cd.anxiety_other ? `（${cd.anxiety_other}）` : "（　　　　　　　　）", { s: F7 })], { af: 10 }));

    // (8) てんかん発作等（けいれん及び意識障害）
    rows.push(p([t("（8）　てんかん発作等（けいれん及び意識障害）", { s: F8, b: true })]));
    if (epi.has_epilepsy) {
      rows.push(p([t(`　　○てんかん発作　発作型（${epi.seizure_type || "イ・ロ・ハ・ニ"}）　頻度（${epi.frequency || "　　"}回／月・年）　最終発作（${epi.last_seizure ? formatDate(epi.last_seizure) : "　　年　　月　　日"}）`, { s: F7 })], { af: 5 }));
    } else {
      rows.push(p([t("　　　てんかん発作　発作型（イ・ロ・ハ・ニ）　頻度（　　回／月・年）　最終発作（　　年　　月　　日）", { s: F7 })], { af: 5 }));
    }
    rows.push(p([t("　　　てんかん発作の型　イ：意識障害はないが、随意運動が失われる発作", { s: 12 })]));
    rows.push(p([t("　　　　　　　　　　　　ロ：意識を失い、行為が途絶するが、倒れない発作", { s: 12 })]));
    rows.push(p([t("　　　　　　　　　　　　ハ：意識障害の有無を問わず、転倒する発作", { s: 12 })]));
    rows.push(p([t("　　　　　　　　　　　　ニ：意識障害を呈し、状況にそぐわない行為を示す発作", { s: 12 })]));
    rows.push(p([t(`　　${epi.consciousness_disorder ? "○意識障害" : "　意識障害"}　${ci(epi.other ? ["その他"] : [], "その他")}`, { s: F7 }),
      t(epi.other ? `（${epi.other}）` : "（　　　　　　　　）", { s: F7 })], { af: 10 }));

    // (9) 精神作用物質の乱用、依存等
    rows.push(p([t("（9）　精神作用物質の乱用、依存等", { s: F8, b: true })]));
    rows.push(p([t(`　　${ci(sub,"アルコール")}　${ci(sub,"覚醒剤")}　${ci(sub,"有機溶剤")}　${ci(sub,"その他")}`, { s: F7 }),
      t(sub.includes("その他") && cd.substance_other ? `（${cd.substance_other}）` : "（　　　　　　　　）", { s: F7 })]));
    rows.push(p([t(`　　　ア乱用　イ依存　ウ残遺性・遅発性精神病性障害　エその他（${cd.substance_detail || "　　　　　　"}）`, { s: F7 })], { af: 10 }));

    // (10) 知能、記憶、学習等の障害
    rows.push(p([t("（10）　知能、記憶、学習等の障害", { s: F8, b: true })]));
    rows.push(p([t(`　　${ci(cog,"知的障害")}（精神遅滞）　ア軽度　イ中等度　ウ重度`, { s: F7 })]));
    rows.push(p([t(`　　${ci(cog,"認知症")}　${ci(cog,"その他の記憶障害")}（${cd.memory_detail || "　　　　　　"}）`, { s: F7 })]));
    rows.push(p([t(`　　${ci(cog,"学習の困難")}　ア読み　イ書き　ウ算数　エその他（${cd.learning_detail || "　　　　　　"}）`, { s: F7 })]));
    rows.push(p([t(`　　${ci(cog,"遂行機能障害")}　${ci(cog,"注意障害")}　${ci(cog,"その他")}`, { s: F7 }),
      t(cog.includes("その他") && cd.cognitive_other ? `（${cd.cognitive_other}）` : "（　　　　　　）", { s: F7 })], { af: 10 }));

    // (11) 広汎性発達障害関連症状
    rows.push(p([t("（11）　広汎性発達障害関連症状", { s: F8, b: true })]));
    rows.push(p([t(`　　${ci(dev,"相互的な社会関係の質的障害")}　${ci(dev,"コミュニケーションのパターンにおける質的障害")}`, { s: F7 })]));
    rows.push(p([t(`　　${ci(dev,"限定した常同的で反復的な関心と活動")}　${ci(dev,"その他")}`, { s: F7 }),
      t(dev.includes("その他") && cd.developmental_other ? `（${cd.developmental_other}）` : "（　　　　　　　　）", { s: F7 })], { af: 10 }));

    // (12) その他
    rows.push(p([t("（12）　その他", { s: F8, b: true }),
      t(`（${cd.other || "　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　"}）`, { s: F7 })]));

    return rows;
  }

  // --- 福祉サービス ---
  function sec7() {
    return [
      p([t("（該当する項目を○で囲んでください。）", { s: F7 })], { af: 10 }),
      p([t(`${ci(wsItems,"自立訓練")}（生活訓練）`, { s: F8 })]),
      p([t(`${ci(wsItems,"共同生活援助")}（グループホーム）`, { s: F8 })]),
      p([t(`${ci(wsItems,"居宅介護")}（ホームヘルプ）`, { s: F8 })]),
      p([t(`${ci(wsItems,"その他の障害福祉サービス")}（${ws.other_detail || "　　　　　　"}）`, { s: F8 })]),
      p([t(`${ci(wsItems,"訪問指導等")}`, { s: F8 })]),
      p([t(`${ci(wsItems,"なし")}`, { s: F8 })]),
    ];
  }

  // --- 訪問看護 ---
  const vn = tr.visiting_nurse || tr.designated_doctor;

  // ========== ドキュメント構築 ==========
  const doc = new Document({
    styles: { default: { document: { run: { font: FONT, size: F8 } } } },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 500, right: 680, bottom: 400, left: 680 },
        },
      },
      children: [
        // 様式番号
        p([t("第12号様式（第８条関係）", { s: F7 })], { a: AlignmentType.RIGHT, af: 20 }),

        // タイトル
        p([t("自立支援医療診断書（精神通院）", { s: F12, b: true })], { a: AlignmentType.CENTER, af: 60 }),

        // ========== 患者情報テーブル ==========
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("氏　　名", 1100),
            c([
              p([t("フリガナ", { s: 10 })]),
              p(pt.name || "", { s: F10 }),
            ], 3600),
            hc("生年月日", 1000),
            c([p([
              t("明・大・昭・平", { s: F7 }),
              t(pt.birthdate ? `　${formatDate(pt.birthdate)}` : "　　年　　月　　日生", { s: F8 }),
              t(pt.age != null ? `（${pt.age}歳）` : "（　　歳）", { s: F8 }),
            ])], 2506),
            hc("性別", 600),
            c([p([t(pt.sex ? (pt.sex === "男" ? "○男・　女" : "　男・○女") : "男・女", { s: F8 })])], TW-1100-3600-1000-2506-600),
          ]}),
          new TableRow({ children: [
            hc("住　　所", 1100),
            c([
              p([t("〒", { s: F7 }), t(pt.postal_code || "", { s: F7 })]),
              p(pt.address || "", { s: F8 }),
            ], TW - 1100, { cs: 5 }),
          ]}),
        ]}),

        p("", { af: 20 }),

        // ========== メインテーブル（１〜８） ==========
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [

          // --- 1 病名 ---
          // ヘッダー注記
          new TableRow({ children: [
            hc("１　病名", 1600, { rs: 3 }),
            c([
              p([t("（ICDコードは、F00～F99、G40のいずれかを記載してください。）", { s: 10, c: "666666" })], { af: 5 }),
              p([t("（1）主たる精神障害　", { s: F8, b: true }), t(data.diagnosis_main || "", { s: F9 })]),
              p([t("　　　　　ICDコード　", { s: F7 }), t(data.icd_code_main || "Ｆ　　", { s: F8, b: true })]),
            ], TW - 1600),
          ]}),
          new TableRow({ children: [
            c([
              p([t("（2）従たる精神障害　", { s: F8, b: true }), t(data.diagnosis_sub || "", { s: F9 })]),
              p([t("　　　　　ICDコード　", { s: F7 }), t(data.icd_code_sub || "", { s: F8 })]),
            ], TW - 1600),
          ]}),
          new TableRow({ children: [
            c([
              p([t("（3）身体合併症　", { s: F8, b: true }), t(data.comorbidity || "", { s: F8 })]),
              p([t("（身体合併症欄は、精神疾患に起因する疾患のみご記載ください。それ以外は８の備考欄へ）", { s: 10, c: "666666" })]),
            ], TW - 1600),
          ]}),

          // --- 2 発病から現在までの病歴 ---
          new TableRow({ children: [
            hc("２　発病から現在\nまでの病歴", 1600),
            c([
              p([t("（推定発病年月、発病状況、初発症状、治療の経過等を記載してください。）", { s: 10, c: "666666" })], { af: 10 }),
              p([t(`（推定発病時期　${data.onset_date ? formatDate(data.onset_date).replace(/日$/, "") + "頃" : "　　年　　月頃"}、　初診日　${data.first_visit_date ? formatDate(data.first_visit_date).replace(/日$/, "") : "　　年　　月"}）`, { s: F7, b: true })], { af: 20 }),
              ...(data.clinical_course || "").split("\n").map(l => p(l, { s: F8, l: 240 })),
            ], TW - 1600),
          ]}),

          // --- 3 現在の病状、状態像等 ---
          new TableRow({ children: [
            hc("３　現在の病状、\n状態像等\n\n（該当する項目を\n○で囲んで\nください。）", 1600),
            c(sec3(), TW - 1600),
          ]}),

          // --- 4 具体的程度 ---
          new TableRow({ children: [
            hc("４　３の病状、\n状態像等の具体的\n程度、病状、\n検査所見等", 1600),
            c([
              p([t("（※現在の病状を日常生活への影響や診察時の様子をふまえて具体的にご記載ください。）", { s: 10, c: "666666" })], { af: 10 }),
              ...(data.condition_detail || "").split("\n").map(l => p(l, { s: F8, l: 240 })),
            ], TW - 1600),
          ]}),

          // --- 5 現在の治療内容 ---
          // (1) 投薬内容
          new TableRow({ children: [
            hc("５　現在の\n治療内容", 1600, { rs: 3 }),
            c([
              p([t("（1）投薬内容", { s: F8, b: true })], { af: 5 }),
              p([t("（※診断名に対する薬剤名（商品名可）をご記載ください。）", { s: 10, c: "666666" })], { af: 10 }),
              ...(medsLines.length > 0 ? medsLines.map(l => p(l, { s: F8, l: 240 })) : [p("", { s: F8 })]),
            ], TW - 1600),
          ]}),
          // (2) 精神療法等
          new TableRow({ children: [
            c([
              p([t("（2）精神療法等", { s: F8, b: true })], { af: 5 }),
              p([t("（※「通院精神療法」「行っている」「支持的精神療法」等の簡単な記載でなく、医学的視点から継続的な通院治療の必要性がわかるように診察時にどのような治療や指導がなされているかをご記載ください。）", { s: 10, c: "666666" })], { af: 10 }),
              ...(tr.psychotherapy || "").split("\n").map(l => p(l, { s: F8, l: 240 })),
            ], TW - 1600),
          ]}),
          // (3) 訪問看護指示の有無
          new TableRow({ children: [
            c([
              p([t("（3）訪問看護指示の有無（　", { s: F8, b: true }),
                t(tr.visiting_nurse === "有" ? "○有" : "有", { s: F8 }),
                t("　・　", { s: F8 }),
                t(tr.visiting_nurse === "無" || (!tr.visiting_nurse) ? "○無" : "無", { s: F8 }),
                t("　）", { s: F8, b: true }),
              ]),
            ], TW - 1600),
          ]}),

          // --- 6 今後の治療方針 ---
          new TableRow({ children: [
            hc("６　今後の\n治療方針", 1600),
            c([
              p([t("（※治療目標をふまえて、継続的に行っていく治療方法をご記載ください。）", { s: 10, c: "666666" })], { af: 10 }),
              ...(data.treatment_plan || "").split("\n").map(l => p(l, { s: F8, l: 240 })),
            ], TW - 1600),
          ]}),

          // --- 7 障害福祉サービス等の利用状況 ---
          new TableRow({ children: [
            hc("７　現在の障害\n福祉サービス等の\n利用状況", 1600),
            c(sec7(), TW - 1600),
          ]}),

          // --- 8 備考 ---
          new TableRow({ children: [
            hc("８　備考", 1600),
            c([
              ...(data.remarks || "").split("\n").map(l => p(l, { s: F8, l: 240 })),
              ...(data.remarks ? [] : [p("", { s: F8 })]),
            ], TW - 1600),
          ]}),
        ]}),

        p("", { af: 40 }),

        // ========== 日付 ==========
        p(data.date ? formatDate(data.date) : "　　年　　月　　日", { a: AlignmentType.RIGHT, af: 30, s: F9 }),

        // ========== 医療機関コード・指定医情報 ==========
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("医療機関コード", 2000),
            c(ins.code || "", 2400),
            c([
              p([t("精神保健指定医の証の番号：", { s: F7 }), t(ins.designated_doctor_number || "", { s: F7 })]),
              p([t("精神医療従事年数：", { s: F7 }), t(ins.doctor_years != null ? `${ins.doctor_years}年` : "　　年", { s: F7 })]),
            ], TW - 2000 - 2400),
          ]}),
        ]}),

        p("", { af: 15 }),

        // ========== 医療機関情報 ==========
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("医療機関所在地", 2000),
            c(ins.address || "", TW - 2000),
          ]}),
          new TableRow({ children: [
            hc("名　　称", 2000),
            c(ins.name || "", TW - 2000),
          ]}),
          new TableRow({ children: [
            hc("電話番号", 2000),
            c(ins.phone || "", TW - 2000),
          ]}),
          new TableRow({ children: [
            hc("医師氏名", 2000),
            c([p([t(ins.doctor || "", { s: F9 }), t("　　　（自筆又は記名捺印）", { s: F7, c: "999999" })])], TW - 2000),
          ]}),
        ]}),

        p("", { af: 40 }),

        // ========== 重度かつ継続（東京都記載欄） ==========
        p([t("※東京都で記載いたしますので、空欄のままでお願い致します。", { s: 10, c: "666666" })], { af: 10 }),
        new Table({ width: { size: TW, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            hc("自立支援医療対象", 2400),
            c("（該当・非該当）", 2200),
            hc("高額治療継続者\n（重度かつ継続）", 2800),
            c(hc_data.applicable != null
              ? (hc_data.applicable ? "（○該当・非該当）" : "（該当・○非該当）")
              : "（該当・非該当）", TW - 2400 - 2200 - 2800),
          ]}),
        ]}),

        // フッター
        p("", { af: 20 }),
        p([t("（日本工業規格Ａ列３番）", { s: 10, c: "999999" })]),
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log(`生成完了: ${outputPath}`);
}

generate().catch(err => { console.error("生成エラー:", err); process.exit(1); });
