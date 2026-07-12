import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const path = process.argv[2];
const input = await FileBlob.load(path);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheets = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 2000 });
console.log(sheets.ndjson);
const matches = await workbook.inspect({
  kind: "match",
  searchTerm: "XXMONB5PCAT|PCATXXMONB5|XXMONB5|PCAT",
  options: { useRegex: true, maxResults: 50 },
  maxChars: 6000,
  summary: "SKU mapping matches",
});
console.log(matches.ndjson);
const row = await workbook.inspect({ kind: "table", range: "Sheet1!A64:N64", include: "values,formulas", tableMaxRows: 2, tableMaxCols: 14, maxChars: 4000 });
console.log(row.ndjson);
