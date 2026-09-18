/**
 * SCRIPT 2: Extract Strategic Execution AOP Critical Data (Office Script / TypeScript)
 * ====================================================================================
 * Location: Run on 'Strategic Execution Dashboard-17Th_sept.xlsx' in SharePoint.
 * Action in Power Automate: "Run script" -> Select Strategic Execution workbook.
 * Returns: JSON string containing extracted strategic KPI rows.
 */

interface StrategicKPIRow {
  strategicImperative: string;
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
  console.log("Starting extraction from Strategic Execution Dashboard...");
  let aopSheet = workbook.getWorksheet("AOP Critical");
  if (!aopSheet) {
    const sheets = workbook.getWorksheets();
    aopSheet = sheets.find(s => s.getName().toLowerCase().includes("aop") || s.getName().toLowerCase().includes("strategic")) || sheets[0];
  }

  if (!aopSheet) {
    return JSON.stringify([]);
  }

  const usedRange = aopSheet.getUsedRange();
  if (!usedRange) return JSON.stringify([]);

  const values = usedRange.getValues();
  if (values.length < 5) return JSON.stringify([]);

  const extractedRows: StrategicKPIRow[] = [];
  let headerRowIndex = 4;

  for (let r = 0; r < Math.min(12, values.length); r++) {
    const rowStr = values[r].map(v => String(v).toLowerCase()).join(" ");
    if (rowStr.includes("jan") || rowStr.includes("feb")) {
      headerRowIndex = r;
      break;
    }
  }

  const headerRow = values[headerRowIndex];
  const monthCols = getMonthColumnIndexes(headerRow);

  let currentImperative = "Customer Focus";

  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const row = values[r];
    const impVal = String(row[0] || "").trim();
    if (impVal && !impVal.startsWith("MoS") && !impVal.startsWith("Strategic")) {
      currentImperative = normalizeImperativeName(impVal);
    }

    const rowTypeVal = String(row[10] || "").trim().toUpperCase();
    if (rowTypeVal === "T" || rowTypeVal === "A") {
      let kpiName = String(row[1] || "").trim();
      if (!kpiName || kpiName.toLowerCase().startsWith("value lever")) continue;

      const kpiDef = String(row[3] || "").trim();
      if (kpiDef.includes("(EQ)") && !kpiName.includes("(EQ)")) {
        kpiName = `${kpiName} (EQ)`;
      } else if (kpiDef.includes("(CS)") && !kpiName.includes("(CS)")) {
        kpiName = `${kpiName} (CS)`;
      }

      const getMonthVal = (mIdx: number): number | string => {
        const cIdx = monthCols[mIdx];
        if (cIdx !== undefined && cIdx < row.length) {
          const v = row[cIdx];
          return v !== null && v !== undefined ? v : "";
        }
        return "";
      };

      extractedRows.push({
        strategicImperative: currentImperative,
        sourceSheet: aopSheet.getName(),
        kpiName: kpiName,
        owner: String(row[2] || "").trim(),
        definition: kpiDef,
        target2026: row[4] !== undefined ? row[4] : "",
        nature: String(row[7] || "").trim(),
        unit: String(row[8] || "").trim(),
        frequency: String(row[9] || "").trim(),
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

  console.log(`Extracted ${extractedRows.length} strategic KPI rows.`);
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
