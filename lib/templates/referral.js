/**
 * 紹介状（診療情報提供書）docx生成
 * Usage: node referral.js <input.json> <output.docx>
 */
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel,
  Header, Footer, PageNumber, PageBreak, TabStopType, TabStopPosition
} = require("docx");

const inputPath = process.argv[2];
const outputPath = process.argv[3];

if (!inputPath || !outputPath) {
  console.error("Usage: node referral.js <input.json> <output.docx>");
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(inputPath, "utf-8"));
const patient = data.patient || {};

// 日付フォーマット
function formatDate(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
}

// ボーダー定義
const noBorder = { style: BorderStyle.NONE, size: 0 };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
const thinBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };

// 薬リスト文字列
function formatMedications(meds) {
  if (!meds || meds.length === 0) return "特になし";
  return meds.map(m => `${m.name} ${m.dosage} ${m.frequency}`).join("\n");
}

// 情報テーブル用のセル
function infoCell(text, width, bold = false) {
  return new TableCell({
    borders: thinBorders,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({
      children: [new TextRun({ text: text || "", font: "Yu Gothic", size: 20, bold })],
    })],
  });
}

// ラベル付きセル（グレー背景）
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

async function generate() {
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: "Yu Gothic", size: 22 } },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }, // 約2cm
        },
      },
      children: [
        // タイトル
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 300 },
          children: [new TextRun({ text: "診療情報提供書", font: "Yu Gothic", size: 36, bold: true })],
        }),

        // 日付
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 200 },
          children: [new TextRun({ text: formatDate(data.date), font: "Yu Gothic", size: 22 })],
        }),

        // 宛先
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: `${data.to_institution || ""} ${data.to_department || ""}`, font: "Yu Gothic", size: 22 })],
        }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({ text: `${data.to_doctor || "御担当医"} 先生　御侍史`, font: "Yu Gothic", size: 22 })],
        }),

        // 差出人情報
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: data.from_institution || "", font: "Yu Gothic", size: 22 })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: data.from_department || "", font: "Yu Gothic", size: 20 })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: data.from_address || "", font: "Yu Gothic", size: 18 })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: `TEL: ${data.from_phone || ""}`, font: "Yu Gothic", size: 18 })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 300 },
          children: [new TextRun({ text: `医師　${data.from_doctor || ""}`, font: "Yu Gothic", size: 22 })],
        }),

        // 患者情報テーブル
        new Table({
          width: { size: 9638, type: WidthType.DXA },
          columnWidths: [1600, 3219, 1600, 3219],
          rows: [
            new TableRow({
              children: [
                labelCell("患者氏名", 1600),
                infoCell(patient.name || "", 3219),
                labelCell("生年月日", 1600),
                infoCell(`${formatDate(patient.birthdate)} (${patient.age || ""}歳)`, 3219),
              ],
            }),
            new TableRow({
              children: [
                labelCell("性別", 1600),
                infoCell(patient.sex || "", 3219),
                labelCell("", 1600),
                infoCell("", 3219),
              ],
            }),
          ],
        }),

        new Paragraph({ spacing: { before: 200 }, children: [] }),

        // 傷病名
        new Table({
          width: { size: 9638, type: WidthType.DXA },
          columnWidths: [1600, 8038],
          rows: [
            new TableRow({
              children: [
                labelCell("傷病名", 1600),
                new TableCell({
                  borders: thinBorders,
                  width: { size: 8038, type: WidthType.DXA },
                  margins: { top: 40, bottom: 40, left: 80, right: 80 },
                  children: (data.diagnosis && data.diagnosis.length > 0)
                    ? data.diagnosis.map(d =>
                        new Paragraph({ children: [new TextRun({ text: d, font: "Yu Gothic", size: 20 })] })
                      )
                    : [new Paragraph({ children: [new TextRun({ text: "（未記載）", font: "Yu Gothic", size: 20 })] })],
                }),
              ],
            }),
          ],
        }),

        new Paragraph({ spacing: { before: 200 }, children: [] }),

        // 症状経過・治療経過
        new Table({
          width: { size: 9638, type: WidthType.DXA },
          columnWidths: [9638],
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: thinBorders,
                  width: { size: 9638, type: WidthType.DXA },
                  shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
                  margins: { top: 40, bottom: 40, left: 80, right: 80 },
                  children: [new Paragraph({
                    children: [new TextRun({ text: "症状経過および治療経過", font: "Yu Gothic", size: 20, bold: true })],
                  })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders: thinBorders,
                  width: { size: 9638, type: WidthType.DXA },
                  margins: { top: 80, bottom: 80, left: 120, right: 120 },
                  children: (data.clinical_course || "").split("\n").map(line =>
                    new Paragraph({
                      spacing: { after: 80 },
                      children: [new TextRun({ text: line, font: "Yu Gothic", size: 20 })],
                    })
                  ),
                }),
              ],
            }),
          ],
        }),

        new Paragraph({ spacing: { before: 200 }, children: [] }),

        // 現在の処方
        new Table({
          width: { size: 9638, type: WidthType.DXA },
          columnWidths: [9638],
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: thinBorders,
                  width: { size: 9638, type: WidthType.DXA },
                  shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
                  margins: { top: 40, bottom: 40, left: 80, right: 80 },
                  children: [new Paragraph({
                    children: [new TextRun({ text: "現在の処方", font: "Yu Gothic", size: 20, bold: true })],
                  })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders: thinBorders,
                  width: { size: 9638, type: WidthType.DXA },
                  margins: { top: 80, bottom: 80, left: 120, right: 120 },
                  children: (data.current_medications || []).map(m =>
                    new Paragraph({
                      spacing: { after: 40 },
                      children: [new TextRun({ text: `${m.name} ${m.dosage} ${m.frequency}`, font: "Yu Gothic", size: 20 })],
                    })
                  ).concat(
                    (!data.current_medications || data.current_medications.length === 0)
                      ? [new Paragraph({ children: [new TextRun({ text: "特になし", font: "Yu Gothic", size: 20 })] })]
                      : []
                  ),
                }),
              ],
            }),
          ],
        }),

        // 検査結果（あれば）
        ...(data.test_results ? [
          new Paragraph({ spacing: { before: 200 }, children: [] }),
          new Table({
            width: { size: 9638, type: WidthType.DXA },
            columnWidths: [9638],
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    borders: thinBorders,
                    width: { size: 9638, type: WidthType.DXA },
                    shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
                    margins: { top: 40, bottom: 40, left: 80, right: 80 },
                    children: [new Paragraph({
                      children: [new TextRun({ text: "検査結果", font: "Yu Gothic", size: 20, bold: true })],
                    })],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    borders: thinBorders,
                    width: { size: 9638, type: WidthType.DXA },
                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                    children: data.test_results.split("\n").map(line =>
                      new Paragraph({
                        spacing: { after: 80 },
                        children: [new TextRun({ text: line, font: "Yu Gothic", size: 20 })],
                      })
                    ),
                  }),
                ],
              }),
            ],
          }),
        ] : []),

        new Paragraph({ spacing: { before: 200 }, children: [] }),

        // 紹介目的
        new Table({
          width: { size: 9638, type: WidthType.DXA },
          columnWidths: [9638],
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: thinBorders,
                  width: { size: 9638, type: WidthType.DXA },
                  shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
                  margins: { top: 40, bottom: 40, left: 80, right: 80 },
                  children: [new Paragraph({
                    children: [new TextRun({ text: "紹介目的", font: "Yu Gothic", size: 20, bold: true })],
                  })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders: thinBorders,
                  width: { size: 9638, type: WidthType.DXA },
                  margins: { top: 80, bottom: 80, left: 120, right: 120 },
                  children: (data.purpose || "").split("\n").map(line =>
                    new Paragraph({
                      spacing: { after: 80 },
                      children: [new TextRun({ text: line, font: "Yu Gothic", size: 20 })],
                    })
                  ),
                }),
              ],
            }),
          ],
        }),

        // 備考（あれば）
        ...(data.remarks ? [
          new Paragraph({ spacing: { before: 200 }, children: [] }),
          new Table({
            width: { size: 9638, type: WidthType.DXA },
            columnWidths: [9638],
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    borders: thinBorders,
                    width: { size: 9638, type: WidthType.DXA },
                    shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
                    margins: { top: 40, bottom: 40, left: 80, right: 80 },
                    children: [new Paragraph({
                      children: [new TextRun({ text: "備考", font: "Yu Gothic", size: 20, bold: true })],
                    })],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    borders: thinBorders,
                    width: { size: 9638, type: WidthType.DXA },
                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                    children: [new Paragraph({
                      children: [new TextRun({ text: data.remarks, font: "Yu Gothic", size: 20 })],
                    })],
                  }),
                ],
              }),
            ],
          }),
        ] : []),
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
