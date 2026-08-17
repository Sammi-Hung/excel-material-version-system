import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "D:/有害物質/outputs/019fa128-90ab-7740-8505-6b6995079dd2";
const outputPath = `${outputDir}/包材調查表_比對測試用_20260730.xlsx`;
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("包材調查表");
sheet.showGridLines = false;

sheet.getRange("A1:I1").merge();
sheet.getRange("A1").values = [["包材調查表 - 比對測試用"]];
sheet.getRange("A2:I2").merge();
sheet.getRange("A2").values = [["使用方式：在系統「調查表上傳 > 包材調查表」填入供應商名稱/代號與品項/包材名稱後，選擇本檔加入包材批次。"]];
sheet.getRange("A3:I3").merge();
sheet.getRange("A3").values = [["建議測試：供應商名稱填「正隆」、品項/包材名稱填「紙箱」可測既有包材修改；也可填新供應商或新品項測右側欄位新增。"]];

sheet.getRange("A7:I8").values = [
  ["No.", "化學物質名稱", "CAS NO.", "等級", "允許濃度", "有 Yes", "含量", "使用用途", "計畫削減"],
  ["", "chemical element", "CAS NO.", "Level", "Threshold Limited(ppm)", "無 No", "content(ppm)", "Purpose", "Hazardous Reduction Plan"],
];

const rows = [
  [1, "Cadmium ( Cd ) / Cadmium Compounds(鎘及化合物)", "7440-43-9", 1, "ND", "無 No", "", "", ""],
  [2, "Lead ( Pb ) / Lead Compounds(鉛及化合物)", "7439-92-1、附表8.1.1", 1, "0.005", "有 Yes", "紙材: 0.005", "Raw material residue", ""],
  [3, "Mercury ( Hg ) / Mercury Compounds(汞及化合物)", "7439-97-6", 1, "ND", "無 No", "", "", ""],
  [4, "Hexavalent-Chromium(Cr6+)Compounds(六價鉻化合物)", "18540-29-9", 1, "ND", "無 No", "", "", ""],
  [5, "PBBs (Polybrominated biphenyls)聚溴聯苯", "059536-65-1", 1, "ND", "無 No", "", "", ""],
  [6, "Polybrominated diphenyl ethers(PBDEs)聚溴二苯醚", "1163-19-5", 1, "ND", "無 No", "", "", ""],
  [7, "氯(Cl) Chlorine", "22537-15-1", 2, "97", "有 Yes", "面紙: 97", "Paper additive", ""],
  [8, "溴(Br) Bromine", "10097-32-2", 2, "ND", "無 No", "", "", ""],
  [9, "Polyvinyl chloride(PVC)and PVC blends聚氯乙烯以及聚氯乙烯混合物", "9002-86-2", 2, "ND", "無 No", "", "", ""],
  [10, "Perfluoroocatane sulfonate(PFOS)全氟辛烷磺酸", "1763-23-1", 2, "ND", "無 No", "", "", ""],
  [11, "Formaldehyde 甲醛", "50-00-0", 2, "<75", "有 Yes", "膠水殘留: 12", "Adhesive residue", ""],
  [12, "測試新增物質 Test Added Substance", "99999-99-9", 2, "0.010", "有 Yes", "0.010", "新增測試用", ""],
];
sheet.getRange("A9:I20").values = rows;

sheet.getRange("A1:I1").format = {
  fill: "#1F6F8B",
  font: { bold: true, color: "#FFFFFF", size: 16 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange("A2:I3").format = {
  fill: "#F3F8FB",
  font: { color: "#244657" },
  wrapText: true,
  verticalAlignment: "center",
};
sheet.getRange("A7:I8").format = {
  fill: "#9CCBEE",
  font: { bold: true },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#000000" },
};
sheet.getRange("A9:I20").format = {
  verticalAlignment: "top",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#B7C3CC" },
};
sheet.getRange("A9:A20").format.horizontalAlignment = "center";
sheet.getRange("D9:D20").format.horizontalAlignment = "center";
sheet.getRange("E9:G20").format.wrapText = true;

sheet.getRange("A1:I20").format.font.name = "Microsoft JhengHei";
sheet.getRange("A1:I20").format.rowHeight = 24;
sheet.getRange("A1:I3").format.rowHeight = 30;
sheet.getRange("A7:I8").format.rowHeight = 32;

sheet.getRange("A:A").format.columnWidth = 7;
sheet.getRange("B:B").format.columnWidth = 48;
sheet.getRange("C:C").format.columnWidth = 26;
sheet.getRange("D:D").format.columnWidth = 10;
sheet.getRange("E:E").format.columnWidth = 20;
sheet.getRange("F:F").format.columnWidth = 14;
sheet.getRange("G:G").format.columnWidth = 22;
sheet.getRange("H:H").format.columnWidth = 22;
sheet.getRange("I:I").format.columnWidth = 28;
sheet.freezePanes.freezeRows(8);

const inspect = await workbook.inspect({
  kind: "table",
  sheetId: "包材調查表",
  range: "A7:I20",
  include: "values",
  tableMaxRows: 14,
  tableMaxCols: 9,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 20 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({ sheetName: "包材調查表", range: "A1:I20", scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/包材調查表_比對測試用_20260730.png`, new Uint8Array(await preview.arrayBuffer()));

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(outputPath);
