import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "D:/有害物質/material_system/data/outputs";
const outputPath = path.join(outputDir, "包材調查表_模擬測試_20260728.xlsx");
const previewPath = path.join(outputDir, "包材調查表_模擬測試_20260728.png");

await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("有害物質調查表");
sheet.showGridLines = false;

sheet.getRange("A1:J1").merge();
sheet.getRange("A1").values = [["有害物質調查表 / Restricted Hazardous Substance survey sheet"]];
sheet.getRange("A3:B3").merge();
sheet.getRange("A3").values = [["供應商Supplier ：模擬包材供應商"]];
sheet.getRange("C3:E3").merge();
sheet.getRange("C3").values = [["品名Item No.： 包材測試-塑膠棧板"]];
sheet.getRange("H3:J3").merge();
sheet.getRange("H3").values = [["公司印："]];
sheet.getRange("A4:B4").merge();
sheet.getRange("A4").values = [["填表人Name：Codex 測試"]];
sheet.getRange("C4:E4").merge();
sheet.getRange("C4").values = [["調查日期 Date：20260728"]];
sheet.getRange("H4:J4").merge();
sheet.getRange("H4").values = [["審核人："]];

sheet.getRange("A6:J7").values = [
  ["No.", "化學物質名稱", "", "等級", "允許濃度", "有Yes", "", "含量", "使用用途", "計畫削減Hazardous Reduction Plan"],
  ["", "chemical element", "CAS NO.", "Level", "Threshold Limited(ppm)", "無No", "", "content(ppm)", "Purpose", ""],
];
sheet.getRange("A8:J8").values = [["Level 1", "", "", "", "", "", "", "", "", ""]];

const rows = [
  [1, "Cadmium ( Cd ) / Cadmium Compounds(鎘及化合物)", "7440-43-9", 1, "ND", "無No", "", "", "", ""],
  [2, "Lead ( Pb ) / Lead Compounds(鉛及化合物)", "7439-92-1", 1, "0.005", "有Yes", "", "0.005", "測試修改允許濃度", ""],
  [3, "Mercury ( Hg ) / Mercury Compounds(汞及化合物)", "7439-97-6", 1, "", "無No", "", "", "", ""],
  [4, "Hexavalent-Chromium(Cr6+)Compounds(六價鉻化合物)", "18540-29-9", 1, "ND", "無No", "", "", "", ""],
  [5, "PBBs (Polybrominated biphenyls)聚溴聯苯", "059536-65-1、附表8.1.5", 1, "ND", "無No", "", "", "", ""],
  [6, "Polybrominated dipphenylethers(PBDEs) / Polybrominated biphenyl ethers(PBBEs)聚溴二苯醚", "1163-19-5\netc.", 1, "15ppm", "有Yes", "", "15ppm", "測試非 No contain", ""],
  [7, "氯(Cl)   Chlorine", "22537-15-1", 2, "<600ppm", "有Yes", "", "420ppm", "測試修改允許濃度", ""],
  [8, "模擬新增包材物質 Test Additive X", "123456-78-9", 2, "25ppm", "有Yes", "", "25ppm", "測試新增細項", ""],
  [9, "模擬新增包材物質 Test Additive X", "123456-78-9", 2, "30ppm", "有Yes", "", "30ppm", "測試同批重複但數值不同", ""],
  [10, "模擬新增包材空白濃度 Blank Package Substance", "98765-43-2", 3, "", "無No", "", "", "測試空白轉 No contain", ""],
];
sheet.getRange("A9:J18").values = rows;

sheet.getRange("A1:J1").format = {
  font: { bold: true, color: "#0F172A", size: 14 },
  fill: "#D9EAF7",
  horizontalAlignment: "center",
};
sheet.getRange("A6:J7").format = {
  fill: "#9DCCF4",
  font: { bold: true, color: "#0F172A" },
  horizontalAlignment: "center",
  verticalAlignment: "middle",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#1F2937" },
};
sheet.getRange("A8:J18").format = {
  borders: { preset: "all", style: "thin", color: "#374151" },
  wrapText: true,
  verticalAlignment: "middle",
};
sheet.getRange("A3:J4").format = {
  borders: { preset: "outside", style: "thin", color: "#CBD5E1" },
  wrapText: true,
};
sheet.getRange("A9:A18").format.horizontalAlignment = "center";
sheet.getRange("D9:F18").format.horizontalAlignment = "center";
sheet.getRange("A:A").format.columnWidth = 8;
sheet.getRange("B:B").format.columnWidth = 48;
sheet.getRange("C:C").format.columnWidth = 24;
sheet.getRange("D:D").format.columnWidth = 10;
sheet.getRange("E:E").format.columnWidth = 22;
sheet.getRange("F:G").format.columnWidth = 10;
sheet.getRange("H:H").format.columnWidth = 18;
sheet.getRange("I:J").format.columnWidth = 24;
sheet.freezePanes.freezeRows(8);

const inspect = await workbook.inspect({
  kind: "table",
  sheetId: "有害物質調查表",
  range: "A6:J18",
  tableMaxRows: 13,
  tableMaxCols: 10,
  maxChars: 5000,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({
  sheetName: "有害物質調查表",
  range: "A1:J18",
  scale: 1,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(JSON.stringify({ outputPath, previewPath }, null, 2));
