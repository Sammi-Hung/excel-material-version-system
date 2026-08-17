import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const outputDir = "D:/有害物質/outputs/019fa128-90ab-7740-8505-6b6995079dd2";
const inputPath = `${outputDir}/包材調查表_比對測試用_20260730.xlsx`;
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const inspect = await workbook.inspect({
  kind: "table",
  sheetId: "包材調查表",
  range: "A7:I20",
  include: "values",
  tableMaxRows: 14,
  tableMaxCols: 9,
});
console.log(inspect.ndjson);

const preview = await workbook.render({ sheetName: "包材調查表", range: "A1:I20", scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/包材調查表_比對測試用_20260730.png`, new Uint8Array(await preview.arrayBuffer()));
