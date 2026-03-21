/**
 * 診断書 docx生成
 * Usage: node diagnosis_certificate.js <input.json> <output.docx>
 */
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType
} = require("docx");

const inputPath = process.argv[2];
const outputPath = process.argv[3];

if (!inputPath || !outputPath) {
  console.error("Usage: node diagnosis_certificate.js <input.json> <output.docx>");
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(inputPath, "utf-8"));

function formatDate(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
}

const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
const thinBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };

function labelCell(text, width) {
  return new TableCell({
    borders: thinBorders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({
      children: [new TextRun({ text, font: "Yu Gothic", size: 20, bold: true })],
    })],
  });
}

function valueCell(text, width) {
  return new TableCell({
    borders: thinBorders,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({
      children: [new TextRun({ text: text || "", font: "Yu Gothic", size: 20 })],
    })],
  });
}

function sectionHeader(text) {
  return new TableRow({
    children: [
      new TableCell({
        borders: thinBorders,
        width: { size: 9638, type: WidthType.DXA },
        columnSpan: 4,
        shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
        margins: { top: 40, bottom: 40, left: 80, right: 80 },
        children: [new Paragraph({
          children: [new TextRun({ text, font: "Yu Gothic", size: 20, bold: true })],
        })],
      }),
    ],
  });
}

function sectionBody(text) {
  return new TableRow({
    children: [
      new TableCell({
        borders: thinBorders,
        width: { size: 9638, type: WidthType.DXA },
        columnSpan: 4,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: (text || "").split("\n").map(line =>
          new Paragraph({
            spacing: { after: 80 },
            children: [new TextRun({ text: line, font: "Yu Gothic", size: 20 })],
          })
        ),
      }),
    ],
  });
}

async function generate() {
  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Yu Gothic", size: 22 } } },
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
        },
      },
      children: [
        // タイトル
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [new TextRun({ text: "診　断　書", font: "Yu Gothic", size: 40, bold: true })],
        }),

        // 日付
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 200 },
          children: [new TextRun({ text: formatDate(data.date), font: "Yu Gothic", size: 22 })],
        }),

        // メインテーブル
        new Table({
          width: { size: 9638, type: WidthType.DXA },
          columnWidths: [1600, 3219, 1600, 3219],
          rows: [
            // 患者情報
            new TableRow({
              children: [
                labelCell("氏名", 1600),
                valueCell(data.patient?.name, 3219),
                labelCell("生年月日", 1600),
                valueCell(`${formatDate(data.patient?.birthdate)} (${data.patient?.age || ""}歳)`, 3219),
              ],
            }),
            new TableRow({
              children: [
                labelCell("性別", 1600),
                valueCell(data.patient?.sex, 3219),
                labelCell("住所", 1600),
                valueCell(data.patient?.address, 3219),
              ],
            }),

            // 病名
            sectionHeader("病名"),
            new TableRow({
              children: [
                new TableCell({
                  borders: thinBorders,
                  width: { size: 9638, type: WidthType.DXA },
                  columnSpan: 4,
                  margins: { top: 60, bottom: 60, left: 120, right: 120 },
                  children: (data.diagnosis || []).map(d =>
                    new Paragraph({ children: [new TextRun({ text: d, font: "Yu Gothic", size: 20 })] })
                  ),
                }),
              ],
            }),

            // 発症日
            new TableRow({
              children: [
                labelCell("発症/初診日", 1600),
                new TableCell({
                  borders: thinBorders,
                  width: { size: 8038, type: WidthType.DXA },
                  columnSpan: 3,
                  margins: { top: 40, bottom: 40, left: 80, right: 80 },
                  children: [new Paragraph({
                    children: [new TextRun({ text: formatDate(data.onset_date), font: "Yu Gothic", size: 20 })],
                  })],
                }),
              ],
            }),

            // 症状の経過
            sectionHeader("症状の経過"),
            sectionBody(data.symptoms),

            // 治療内容
            sectionHeader("治療内容"),
            sectionBody(data.treatment),

            // 今後の見通し
            sectionHeader("今後の見通し"),
            sectionBody(data.prognosis),

            // 就労に関する意見
            ...(data.work_restriction ? [
              sectionHeader("就労に関する意見"),
              new TableRow({
                children: [
                  labelCell("休業の必要性", 1600),
                  valueCell(data.work_restriction.needs_rest ? "あり" : "なし", 3219),
                  labelCell("休業期間", 1600),
                  valueCell(data.work_restriction.rest_period, 3219),
                ],
              }),
              ...(data.work_restriction.restrictions ? [
                new TableRow({
                  children: [
                    labelCell("制限事項", 1600),
                    new TableCell({
                      borders: thinBorders,
                      width: { size: 8038, type: WidthType.DXA },
                      columnSpan: 3,
                      margins: { top: 40, bottom: 40, left: 80, right: 80 },
                      children: [new Paragraph({
                        children: [new TextRun({ text: data.work_restriction.restrictions, font: "Yu Gothic", size: 20 })],
                      })],
                    }),
                  ],
                }),
              ] : []),
            ] : []),

            // 備考
            ...(data.remarks ? [
              sectionHeader("備考"),
              sectionBody(data.remarks),
            ] : []),
          ],
        }),

        new Paragraph({ spacing: { before: 400 }, children: [] }),

        // 用途
        new Paragraph({
          spacing: { after: 400 },
          children: [new TextRun({ text: `上記の通り診断する。（${data.purpose || ""}）`, font: "Yu Gothic", size: 22 })],
        }),

        // 医療機関情報
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: data.institution || "", font: "Yu Gothic", size: 22 })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: data.department || "", font: "Yu Gothic", size: 20 })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: data.institution_address || "", font: "Yu Gothic", size: 18 })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: `TEL: ${data.institution_phone || ""}`, font: "Yu Gothic", size: 18 })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 200 },
          children: [
            new TextRun({ text: "医師　", font: "Yu Gothic", size: 22 }),
            new TextRun({ text: data.doctor || "", font: "Yu Gothic", size: 22 }),
            new TextRun({ text: "　　　　印", font: "Yu Gothic", size: 22, color: "CCCCCC" }),
          ],
        }),
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log(`生成完了: ${outputPath}`);
}

generate().catch(err => {
  console.error("生成エラー:", err);
  process.exit(1);
});
