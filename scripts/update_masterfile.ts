/**
 * SCRIPT 3: Populate Masterfile Worksheets (Office Script / TypeScript)
 * =====================================================================
 * Location: Run on 'Masterfile_DMB_Dashboard.xlsx' in SharePoint.
 * Action in Power Automate: "Run script" -> Select Masterfile workbook.
 * Parameters:
 *   - functionalJson (string): JSON output from SCRIPT 1 (extract_functional_dmb)
 *   - strategicJson (string): JSON output from SCRIPT 2 (extract_strategic_execution)
 */

interface KPIRow {
  functionName?: string;
  strategicImperative?: string;
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

function main(workbook: ExcelScript.Workbook, functionalJson?: string, strategicJson?: string): { dmbRows: number; mprRows: number; status: string } {
  console.log("Starting Masterfile update script...");

  let dmbRowsCount = 0;
  let mprRowsCount = 0;

  // 1. Process Functional Review Data into 'DMB Masterfile'
  if (functionalJson && functionalJson.trim().length > 2) {
    try {
      const functionalRows: KPIRow[] = JSON.parse(functionalJson);
      let dmbSheet = workbook.getWorksheet("DMB Masterfile");
      if (!dmbSheet) {
        dmbSheet = workbook.addWorksheet("DMB Masterfile");
      }
      dmbRowsCount = writeWorksheetData(dmbSheet, "Function", functionalRows, "#002D4B");
      console.log(`Successfully wrote ${dmbRowsCount} rows to DMB Masterfile.`);
    } catch (err) {
      console.error("Error parsing/writing functionalJson:", err);
    }
  }

  // 2. Process Strategic Execution Data into 'MPR Masterfile'
  if (strategicJson && strategicJson.trim().length > 2) {
    try {
      const strategicRows: KPIRow[] = JSON.parse(strategicJson);
      let mprSheet = workbook.getWorksheet("MPR Masterfile");
      if (!mprSheet) {
        mprSheet = workbook.addWorksheet("MPR Masterfile");
      }
      mprRowsCount = writeWorksheetData(mprSheet, "Strategic Imperative", strategicRows, "#0B557F");
      console.log(`Successfully wrote ${mprRowsCount} rows to MPR Masterfile.`);
    } catch (err) {
      console.error("Error parsing/writing strategicJson:", err);
    }
  }

  return {
    dmbRows: dmbRowsCount,
    mprRows: mprRowsCount,
    status: `Masterfile successfully synced: ${dmbRowsCount} DMB rows and ${mprRowsCount} MPR rows updated.`
  };
}

function writeWorksheetData(sheet: ExcelScript.Worksheet, firstColName: string, rows: KPIRow[], headerColor: string): number {
  if (rows.length === 0) return 0;

  // Clear existing content
  sheet.getUsedRange()?.clear();

  const headers = [
    firstColName,
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
    const category = r.functionName || r.strategicImperative || "";
    tableData.push([
      category,
      r.kpiName,
      r.owner,
      r.definition,
      r.target2026,
      r.nature,
      r.unit,
      r.frequency,
      r.rowType,
      r.jan,
      r.feb,
      r.mar,
      r.apr,
      r.may,
      r.jun,
      r.jul,
      r.aug,
      r.sep,
      r.oct,
      r.nov,
      r.dec
    ]);
  }

  const range = sheet.getRangeByIndexes(0, 0, tableData.length, headers.length);
  range.setValues(tableData);

  // Style Header Row
  const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
  headerRange.getFormat().getFill().setColor(headerColor);
  headerRange.getFormat().getFont().setColor("#FFFFFF");
  headerRange.getFormat().getFont().setBold(true);

  // Auto-fit columns
  sheet.getUsedRange()?.getFormat().autofitColumns();

  return tableData.length - 1;
}
