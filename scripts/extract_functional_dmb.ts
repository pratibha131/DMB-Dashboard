/**
 * SCRIPT 1: Extract Functional DMB Review Data (Office Script / TypeScript)
 * =========================================================================
 * Location: Run on 'Functional DMB Review Sheets-17th_sept.xlsx' in SharePoint.
 * Action in Power Automate: "Run script" -> Select Functional DMB workbook.
 * Returns: JSON string containing extracted functional KPI rows.
 */

interface FunctionalKPIRow {
  functionName: string;
  sourceSheet: string;
  kpiName: string;
  owner: string;
  definition: string;
  target2026: number | string;
  nature: string;
  unit: string;
  frequency: string;
  rowType: "T" | "A";
  jan: number | string;
  feb: number | string;
  mar: number | string;
  apr: number | string;
  may: number | string;
  jun: number | string;
  jul: number | string;
  aug: number | string;
  sep: number | string;
  oct: number | string;
  nov: number | string;
  dec: number | string;
}

function main(workbook: ExcelScript.Workbook): string {
  console.log("Starting extraction from Functional DMB Review Sheets...");
  const worksheets = workbook.getWorksheets();
  const functionalKeywords = [
    "quality", "regulatory", "isc", "procurement",
    "r&d", "marketing", "customer service", "commercial excellence",
    "nar", "europe", "growth", "finance"
  ];

  const extractedRows: FunctionalKPIRow[] = [];

  for (const sheet of worksheets) {
    const sheetName = sheet.getName().trim();
    const sheetNameLower = sheetName.toLowerCase();

    // Only process functional review worksheets
    const isFunctional = functionalKeywords.some(keyword => sheetNameLower.includes(keyword));
    if (!isFunctional || sheetNameLower.includes("mastersheet") || sheetNameLower.includes("mpr")) {
      continue;
    }

    const functionName = cleanFunctionName(sheetName);
    const usedRange = sheet.getUsedRange();
    if (!usedRange) continue;

    const values = usedRange.getValues();
    if (values.length < 5) continue;

    // Locate header row containing calendar months
    let headerRowIndex = 4;
    for (let r = 0; r < Math.min(12, values.length); r++) {
      const rowStr = values[r].map(v => String(v).toLowerCase()).join(" ");
      if (rowStr.includes("jan") || rowStr.includes("feb") || rowStr.includes("q1")) {
        headerRowIndex = r;
        break;
      }
    }

    const headerRow = values[headerRowIndex];
    const monthCols = getMonthColumnIndexes(headerRow);

    let currentKpiName = "";
    let currentOwner = "";
    let currentDefinition = "";
    let currentTarget2026: number | string = "";
    let currentNature = "";
    let currentUnit = "";
    let currentFreq = "";

    for (let r = headerRowIndex + 1; r < values.length; r++) {
      const row = values[r];
      const rowTypeVal = String(row[10] || "").trim().toUpperCase();

      if (rowTypeVal === "T" || rowTypeVal === "A") {
        const kpiNameCell = String(row[1] || "").trim();
        if (kpiNameCell && !kpiNameCell.toLowerCase().startsWith("value lever")) {
          currentKpiName = kpiNameCell;
          currentOwner = String(row[2] || "").trim();
          currentDefinition = String(row[3] || "").trim();
          currentTarget2026 = row[4] !== undefined && row[4] !== null ? row[4] : "";
          currentNature = String(row[7] || "").trim();
          currentUnit = String(row[8] || "").trim();
          currentFreq = String(row[9] || "").trim();
        }

        if (!currentKpiName) continue;

        const getMonthVal = (mIdx: number): number | string => {
          const cIdx = monthCols[mIdx];
          if (cIdx !== undefined && cIdx < row.length) {
            const v = row[cIdx];
            return v !== null && v !== undefined ? v : "";
          }
          return "";
        };

        extractedRows.push({
          functionName: functionName,
          sourceSheet: sheetName,
          kpiName: currentKpiName,
          owner: currentOwner,
          definition: currentDefinition,
          target2026: currentTarget2026,
          nature: currentNature,
          unit: currentUnit,
          frequency: currentFreq,
          rowType: rowTypeVal as "T" | "A",
          jan: getMonthVal(0),
          feb: getMonthVal(1),
          mar: getMonthVal(2),
          apr: getMonthVal(3),
          may: getMonthVal(4),
          jun: getMonthVal(5),
          jul: getMonthVal(6),
          aug: getMonthVal(7),
          sep: getMonthVal(8),
          oct: getMonthVal(9),
          nov: getMonthVal(10),
          dec: getMonthVal(11)
        });
      }
    }
  }

  console.log(`Extracted ${extractedRows.length} functional KPI rows.`);
  return JSON.stringify(extractedRows);
}

function getMonthColumnIndexes(headerRow: (string | number | boolean)[]): { [monthIndex: number]: number } {
  const prefixes = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const mapping: { [monthIndex: number]: number } = {};
  for (let m = 0; m < prefixes.length; m++) {
    const p = prefixes[m];
    for (let c = 0; c < headerRow.length; c++) {
      const val = String(headerRow[c] || "").trim().toLowerCase();
      if (val.startsWith(p)) {
        mapping[m] = c;
        break;
      }
    }
  }
  return mapping;
}

function cleanFunctionName(sheetName: string): string {
  const clean = sheetName.replace(/^[0-9]+[.\s]*/, "").trim();
  const lower = clean.toLowerCase();
  if (lower.includes("customer")) return "Customer Service";
  if (lower.includes("quality")) return "Quality";
  if (lower.includes("regulatory")) return "Regulatory";
  if (lower.includes("isc")) return "ISC & Procurement";
  if (lower.includes("procurement")) return "ISC & Procurement";
  if (lower.includes("marketing")) return "Marketing";
  if (lower.includes("r&d")) return "R&D";
  if (lower.includes("commercial")) return "Commercial Excellence";
  return clean;
}
