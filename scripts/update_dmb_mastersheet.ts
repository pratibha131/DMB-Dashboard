/**
 * DMB Dashboard Masterfile Consolidator — Excel Office Script (TypeScript)
 * =========================================================================
 * Purpose:
 *   Automates extracting KPI Target/Actual values from Functional DMB Review Sheets
 *   and Strategic Execution Dashboard, and populates the canonical Masterfile_DMB_Dashboard.xlsx.
 *
 * Compatible with:
 *   - Excel Online (Automate > Office Scripts)
 *   - Microsoft Power Automate ("Run script" action in cloud flow)
 */

interface KPIRowData {
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
  monthlyValues: (number | string)[]; // 12 elements (Jan to Dec)
}

interface ProcessedKPI {
  functionName: string;
  kpiName: string;
  owner: string;
  definition: string;
  target2026: number | string;
  nature: string;
  unit: string;
  frequency: string;
  targetMonthly: (number | string)[];
  actualMonthly: (number | string)[];
}

/**
 * Main Office Script entry point executed by Power Automate.
 */
function main(workbook: ExcelScript.Workbook): { success: boolean; rowsProcessed: number; message: string } {
  console.log("Starting DMB Masterfile consolidation script...");

  // 1. Identify or Create 'DMB Masterfile' Sheet
  let dmbMasterSheet = workbook.getWorksheet("DMB Masterfile");
  if (!dmbMasterSheet) {
    dmbMasterSheet = workbook.addWorksheet("DMB Masterfile");
  }

  // 2. Identify or Create 'MPR Masterfile' Sheet
  let mprMasterSheet = workbook.getWorksheet("MPR Masterfile");
  if (!mprMasterSheet) {
    mprMasterSheet = workbook.addWorksheet("MPR Masterfile");
  }

  const worksheets = workbook.getWorksheets();
  const functionalKPIs: KPIRowData[] = [];
  const strategicKPIs: KPIRowData[] = [];

  const functionalSheetKeywords = [
    "quality",
    "regulatory",
    "isc",
    "procurement",
    "r&d",
    "marketing",
    "customer service",
    "commercial excellence",
    "nar",
    "europe",
    "growth",
    "finance"
  ];

  for (const sheet of worksheets) {
    const sheetName = sheet.getName().trim();
    const sheetNameLower = sheetName.toLowerCase();

    // Skip output master sheets
    if (sheetName === "DMB Masterfile" || sheetName === "MPR Masterfile" || sheetName === "MasterSheet") {
      continue;
    }

    // Check if it's Strategic Execution (AOP Critical)
    if (sheetNameLower.includes("aop critical") || sheetNameLower.includes("strategic")) {
      const rows = extractStrategicKPIs(sheet);
      strategicKPIs.push(...rows);
      continue;
    }

    // Check if it's a Functional Review sheet
    const isFunctional = functionalSheetKeywords.some(keyword => sheetNameLower.includes(keyword));
    if (isFunctional) {
      const rows = extractFunctionalSheetKPIs(sheet);
      functionalKPIs.push(...rows);
    }
  }

  console.log(`Extracted ${functionalKPIs.length} functional rows and ${strategicKPIs.length} strategic rows.`);

  // 3. Write Consolidated Data to DMB Masterfile
  const dmbRowsWritten = writeDMBMasterfile(dmbMasterSheet, functionalKPIs);

  // 4. Write Consolidated Data to MPR Masterfile
  const mprRowsWritten = writeMPRMasterfile(mprMasterSheet, strategicKPIs);

  return {
    success: true,
    rowsProcessed: dmbRowsWritten + mprRowsWritten,
    message: `Consolidated ${dmbRowsWritten} DMB rows and ${mprRowsWritten} MPR rows successfully.`
  };
}

/**
 * Extract Target and Actual KPI rows from a functional review worksheet.
 */
function extractFunctionalSheetKPIs(sheet: ExcelScript.Worksheet): KPIRowData[] {
  const sheetName = sheet.getName().trim();
  const functionName = cleanFunctionName(sheetName);
  const usedRange = sheet.getUsedRange();
  if (!usedRange) return [];

  const values = usedRange.getValues();
  if (values.length < 5) return [];

  const rows: KPIRowData[] = [];

  // Locate header row containing 'Jan', 'Feb', etc.
  let headerRowIndex = -1;
  for (let r = 0; r < Math.min(12, values.length); r++) {
    const rowStr = values[r].map(v => String(v).toLowerCase()).join(" ");
    if (rowStr.includes("jan") || rowStr.includes("feb") || rowStr.includes("q1")) {
      headerRowIndex = r;
      break;
    }
  }

  if (headerRowIndex === -1) headerRowIndex = 4;

  const headerRow = values[headerRowIndex];
  const monthColIndexes = getMonthColumnIndexes(headerRow);

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

      const monthlyValues: (number | string)[] = [];
      for (let m = 0; m < 12; m++) {
        const colIdx = monthColIndexes[m];
        const val = colIdx !== undefined && colIdx < row.length ? row[colIdx] : "";
        monthlyValues.push(val !== null && val !== undefined ? val : "");
      }

      rows.push({
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
        monthlyValues: monthlyValues
      });
    }
  }

  return rows;
}

/**
 * Extract Strategic Imperative KPIs from the AOP Critical worksheet.
 */
function extractStrategicKPIs(sheet: ExcelScript.Worksheet): KPIRowData[] {
  const usedRange = sheet.getUsedRange();
  if (!usedRange) return [];

  const values = usedRange.getValues();
  if (values.length < 5) return [];

  const rows: KPIRowData[] = [];
  let headerRowIndex = 4;

  for (let r = 0; r < Math.min(12, values.length); r++) {
    const rowStr = values[r].map(v => String(v).toLowerCase()).join(" ");
    if (rowStr.includes("jan") || rowStr.includes("feb")) {
      headerRowIndex = r;
      break;
    }
  }

  const headerRow = values[headerRowIndex];
  const monthColIndexes = getMonthColumnIndexes(headerRow);

  let currentImperative = "Strategic Imperative";

  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const row = values[r];
    const impVal = String(row[0] || "").trim();
    if (impVal && !impVal.startsWith("MoS") && !impVal.startsWith("Strategic")) {
      currentImperative = normalizeImperativeName(impVal);
    }

    const rowTypeVal = String(row[10] || "").trim().toUpperCase();
    if (rowTypeVal === "T" || rowTypeVal === "A") {
      const kpiName = String(row[1] || "").trim();
      if (!kpiName || kpiName.toLowerCase().startsWith("value lever")) continue;

      const monthlyValues: (number | string)[] = [];
      for (let m = 0; m < 12; m++) {
        const colIdx = monthColIndexes[m];
        const val = colIdx !== undefined && colIdx < row.length ? row[colIdx] : "";
        monthlyValues.push(val !== null && val !== undefined ? val : "");
      }

      rows.push({
        functionName: currentImperative,
        sourceSheet: sheet.getName(),
        kpiName: kpiName,
        owner: String(row[2] || "").trim(),
        definition: String(row[3] || "").trim(),
        target2026: row[4] !== undefined ? row[4] : "",
        nature: String(row[7] || "").trim(),
        unit: String(row[8] || "").trim(),
        frequency: String(row[9] || "").trim(),
        rowType: rowTypeVal as "T" | "A",
        monthlyValues: monthlyValues
      });
    }
  }

  return rows;
}

/**
 * Write structured rows into DMB Masterfile.
 */
function writeDMBMasterfile(sheet: ExcelScript.Worksheet, rows: KPIRowData[]): number {
  sheet.getUsedRange()?.clear();

  const headers = [
    "Function",
    "KPI Name",
    "Owner",
    "Definition",
    "2026 Target",
    "Nature",
    "Unit",
    "Frequency",
    "Type",
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec"
  ];

  const tableData: (string | number)[][] = [headers];

  for (const r of rows) {
    tableData.push([
      r.functionName,
      r.kpiName,
      r.owner,
      r.definition,
      r.target2026,
      r.nature,
      r.unit,
      r.frequency,
      r.rowType,
      ...r.monthlyValues
    ]);
  }

  const range = sheet.getRangeByIndexes(0, 0, tableData.length, headers.length);
  range.setValues(tableData);

  // Format Header Row
  const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
  headerRange.getFormat().getFill().setColor("#002D4B");
  headerRange.getFormat().getFont().setColor("#FFFFFF");
  headerRange.getFormat().getFont().setBold(true);

  return tableData.length - 1;
}

/**
 * Write structured rows into MPR Masterfile.
 */
function writeMPRMasterfile(sheet: ExcelScript.Worksheet, rows: KPIRowData[]): number {
  sheet.getUsedRange()?.clear();

  const headers = [
    "Strategic Imperative",
    "KPI Name",
    "Owner",
    "Definition",
    "2026 Target",
    "Nature",
    "Unit",
    "Frequency",
    "Type",
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec"
  ];

  const tableData: (string | number)[][] = [headers];

  for (const r of rows) {
    tableData.push([
      r.functionName,
      r.kpiName,
      r.owner,
      r.definition,
      r.target2026,
      r.nature,
      r.unit,
      r.frequency,
      r.rowType,
      ...r.monthlyValues
    ]);
  }

  const range = sheet.getRangeByIndexes(0, 0, tableData.length, headers.length);
  range.setValues(tableData);

  // Format Header Row
  const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
  headerRange.getFormat().getFill().setColor("#0B557F");
  headerRange.getFormat().getFont().setColor("#FFFFFF");
  headerRange.getFormat().getFont().setBold(true);

  return tableData.length - 1;
}

/**
 * Helper: Map 12 calendar months to matching column index in row header.
 */
function getMonthColumnIndexes(headerRow: (string | number | boolean)[]): { [monthIndex: number]: number } {
  const prefixes = [
    "jan", "feb", "mar", "apr", "may", "jun",
    "jul", "aug", "sep", "oct", "nov", "dec"
  ];

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

/**
 * Clean sheet name to standardized functional name.
 */
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
  if (lower.includes("finance")) return "Finance";
  return clean;
}

/**
 * Normalize Strategic Imperative names.
 */
function normalizeImperativeName(raw: string): string {
  const lower = raw.toLowerCase();
  if (lower.includes("customer")) return "Customer Focus";
  if (lower.includes("deliverability") || lower.includes("profitability")) return "Improve Deliverability & Profitability";
  if (lower.includes("safety") || lower.includes("quality")) return "Patient Safety & Quality";
  if (lower.includes("commercial") || lower.includes("growth")) return "Drive growth through Commercial Excellence";
  if (lower.includes("roadmap") || lower.includes("innovation")) return "Roadmap competitiveness & Innovation agility";
  if (lower.includes("esg") || lower.includes("environmental") || lower.includes("governance")) return "Environmental, Social & Governance (ESG)";
  return raw.trim();
}
