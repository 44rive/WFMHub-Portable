/**
 * Governed writer for the unchanged FTE Count.xlsx contract.
 *
 * Install in Excel on the web and call it only from the serialized
 * WFM_Sync_Absence_To_FTE flow. The script never touches tblFTEAgents.
 */

type RequestType = "PTO" | "AWAY";
type WriterOperation = "UPSERT" | "CANCEL";
type CellValue = string | number | boolean;

interface AbsenceCommand {
  contractVersion: "1.0.0";
  requestId: string;
  operation: WriterOperation;
  requestType: RequestType;
  clientId: string;
  employeeName: string;
  startDate: string;
  endDate?: string;
  dayCoverage?: "Full day" | "Partial day";
  startTime?: string;
  endTime?: string;
  awayType?: "Long sickness" | "Maternity/Parental" | "Administrative leave" | "Other";
  comment?: string;
  today: string;
}

interface WriterResult {
  ok: boolean;
  requestId: string;
  action: "INSERTED" | "UPDATED" | "NO_CHANGE" | "ERROR";
  tableName: string;
  workbookStatus: string;
  matchCount: number;
  message: string;
}

const PTO_HEADERS = [
  "Client ID", "Name", "Start date", "End date", "Day coverage",
  "Start time", "End time", "PTO type", "Approval status", "Comment"
];

const AWAY_HEADERS = [
  "Client ID", "Name", "Start date", "End date", "Away type",
  "Case status", "Comment"
];

function cleanText(value: string | undefined): string {
  return String(value ?? "").trim();
}

function upper(value: string | undefined): string {
  return cleanText(value).toUpperCase();
}

function requireIsoDate(value: string, label: string): string {
  const candidate = cleanText(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(candidate)) {
    throw new Error(`${label} must use yyyy-MM-dd.`);
  }
  const parsed = new Date(`${candidate}T00:00:00Z`);
  if (Number.isNaN(parsed.getTime()) || parsed.toISOString().slice(0, 10) !== candidate) {
    throw new Error(`${label} is not a valid calendar date.`);
  }
  return candidate;
}

function requireTime(value: string, label: string): string {
  const candidate = cleanText(value);
  const match = /^(\d{2}):(\d{2})(?::(\d{2}))?$/.exec(candidate);
  if (!match) {
    throw new Error(`${label} must use HH:mm:ss.`);
  }
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  const second = Number(match[3] ?? "0");
  if (hour > 23 || minute > 59 || second > 59) {
    throw new Error(`${label} is not a valid time.`);
  }
  return `${match[1]}:${match[2]}:${String(second).padStart(2, "0")}`;
}

function timeToSeconds(value: string): number {
  const parts = value.split(":").map(Number);
  return parts[0] * 3600 + parts[1] * 60 + parts[2];
}

function dateFromCell(value: CellValue): string {
  if (typeof value === "number") {
    const milliseconds = Date.UTC(1899, 11, 30) + Math.round(value * 86400000);
    return new Date(milliseconds).toISOString().slice(0, 10);
  }
  const text = cleanText(String(value));
  if (!text) return "";
  const iso = /^(\d{4}-\d{2}-\d{2})/.exec(text);
  if (iso) return iso[1];
  const dmy = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/.exec(text);
  if (dmy) {
    return `${dmy[3]}-${dmy[2].padStart(2, "0")}-${dmy[1].padStart(2, "0")}`;
  }
  return text;
}

function timeFromCell(value: CellValue): string {
  if (typeof value === "number") {
    let seconds = Math.round((value % 1) * 86400);
    if (seconds === 86400) seconds = 0;
    const hour = Math.floor(seconds / 3600);
    const minute = Math.floor((seconds % 3600) / 60);
    const second = seconds % 60;
    return `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}:${String(second).padStart(2, "0")}`;
  }
  const text = cleanText(String(value));
  if (!text) return "";
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(text);
  if (!match) return text;
  return `${match[1].padStart(2, "0")}:${match[2]}:${match[3] ?? "00"}`;
}

function getTable(workbook: ExcelScript.Workbook, tableName: string): ExcelScript.Table {
  const table = workbook.getTables().find(item => item.getName() === tableName);
  if (!table) throw new Error(`Required Excel table ${tableName} was not found.`);
  return table;
}

function getHeaderIndex(table: ExcelScript.Table, required: string[]): Record<string, number> {
  const headers = table.getHeaderRowRange().getTexts()[0].map(item => item.trim());
  const index: Record<string, number> = {};
  headers.forEach((header, position) => { index[header] = position; });
  const missing = required.filter(header => index[header] === undefined);
  if (missing.length > 0) {
    throw new Error(`Table ${table.getName()} is missing columns: ${missing.join(", ")}.`);
  }
  return index;
}

function deriveAwayStatus(command: AbsenceCommand): string {
  if (command.operation === "CANCEL") return "Cancelled";
  const today = requireIsoDate(command.today, "today");
  const start = requireIsoDate(command.startDate, "Start date");
  const end = cleanText(command.endDate);
  // The current WFMHub contract requires an open-ended Away row to be Active.
  // Its future Start date still prevents it from applying before that date.
  if (!end) return "Active";
  const validEnd = requireIsoDate(end, "End date");
  if (validEnd < today) return "Closed";
  if (start > today) return "Planned";
  return "Active";
}

function validateCommand(command: AbsenceCommand): void {
  if (command.contractVersion !== "1.0.0") throw new Error("Unsupported contract version.");
  if (!cleanText(command.requestId)) throw new Error("Request ID is required.");
  if (!cleanText(command.clientId)) throw new Error("Client ID is required.");
  if (!cleanText(command.employeeName)) throw new Error("Employee name is required.");
  const start = requireIsoDate(command.startDate, "Start date");
  const end = cleanText(command.endDate);
  if (end && requireIsoDate(end, "End date") < start) {
    throw new Error("End date cannot be earlier than Start date.");
  }
  if (command.requestType === "PTO") {
    if (!end) throw new Error("PTO End date is required.");
    if (command.dayCoverage !== "Full day" && command.dayCoverage !== "Partial day") {
      throw new Error("PTO Day coverage must be Full day or Partial day.");
    }
    if (command.dayCoverage === "Partial day") {
      if (end !== start) throw new Error("Partial-day PTO must start and end on the same date.");
      const startTime = requireTime(cleanText(command.startTime), "Start time");
      const endTime = requireTime(cleanText(command.endTime), "End time");
      if (timeToSeconds(endTime) <= timeToSeconds(startTime)) {
        throw new Error("End time must be later than Start time.");
      }
    }
  } else if (command.requestType === "AWAY") {
    if (!["Long sickness", "Maternity/Parental", "Administrative leave", "Other"].includes(cleanText(command.awayType))) {
      throw new Error("A valid Away type is required.");
    }
  } else {
    throw new Error("Request type must be PTO or AWAY.");
  }
}

function buildPtoRow(
  headers: string[], index: Record<string, number>, command: AbsenceCommand,
  existing?: CellValue[]
): CellValue[] {
  const row: CellValue[] = existing ? existing.slice() : headers.map(() => "");
  row[index["Client ID"]] = cleanText(command.clientId);
  row[index["Name"]] = cleanText(command.employeeName);
  row[index["Start date"]] = requireIsoDate(command.startDate, "Start date");
  row[index["End date"]] = requireIsoDate(cleanText(command.endDate), "End date");
  row[index["Day coverage"]] = command.dayCoverage ?? "";
  row[index["Start time"]] = command.dayCoverage === "Partial day" ? requireTime(cleanText(command.startTime), "Start time") : "";
  row[index["End time"]] = command.dayCoverage === "Partial day" ? requireTime(cleanText(command.endTime), "End time") : "";
  row[index["PTO type"]] = "PTO";
  row[index["Approval status"]] = command.operation === "CANCEL" ? "Cancelled" : "Approved";
  row[index["Comment"]] = cleanText(command.comment);
  return row;
}

function buildAwayRow(
  headers: string[], index: Record<string, number>, command: AbsenceCommand,
  existing?: CellValue[]
): CellValue[] {
  const row: CellValue[] = existing ? existing.slice() : headers.map(() => "");
  row[index["Client ID"]] = cleanText(command.clientId);
  row[index["Name"]] = cleanText(command.employeeName);
  row[index["Start date"]] = requireIsoDate(command.startDate, "Start date");
  row[index["End date"]] = cleanText(command.endDate) ? requireIsoDate(cleanText(command.endDate), "End date") : "";
  row[index["Away type"]] = cleanText(command.awayType);
  row[index["Case status"]] = deriveAwayStatus(command);
  row[index["Comment"]] = cleanText(command.comment);
  return row;
}

function applyFormats(rowRange: ExcelScript.Range, index: Record<string, number>, requestType: RequestType): void {
  rowRange.getCell(0, index["Client ID"]).setNumberFormat([["@"]]);
  rowRange.getCell(0, index["Start date"]).setNumberFormat([["yyyy-mm-dd"]]);
  rowRange.getCell(0, index["End date"]).setNumberFormat([["yyyy-mm-dd"]]);
  if (requestType === "PTO") {
    rowRange.getCell(0, index["Start time"]).setNumberFormat([["hh:mm:ss"]]);
    rowRange.getCell(0, index["End time"]).setNumberFormat([["hh:mm:ss"]]);
  }
}

function main(workbook: ExcelScript.Workbook, payloadJson: string): WriterResult {
  let requestId = "";
  let tableName = "";
  try {
    const command = JSON.parse(payloadJson) as AbsenceCommand;
    requestId = cleanText(command.requestId);
    validateCommand(command);
    tableName = command.requestType === "PTO" ? "tblFTEPTO" : "tblFTEAway";
    const required = command.requestType === "PTO" ? PTO_HEADERS : AWAY_HEADERS;
    const table = getTable(workbook, tableName);
    const headers = table.getHeaderRowRange().getTexts()[0].map(item => item.trim());
    const index = getHeaderIndex(table, required);
    const rows: CellValue[][] = table.getRowCount() > 0
      ? table.getRangeBetweenHeaderAndTotal().getValues()
      : [];

    const clientId = cleanText(command.clientId);
    const startDate = requireIsoDate(command.startDate, "Start date");
    const endDate = cleanText(command.endDate);
    const matches: number[] = [];
    rows.forEach((row, position) => {
      const commonMatch = cleanText(String(row[index["Client ID"]])) === clientId
        && dateFromCell(row[index["Start date"]]) === startDate
        && dateFromCell(row[index["End date"]]) === endDate;
      if (!commonMatch) return;
      if (command.requestType === "PTO") {
        const coverageMatch = upper(String(row[index["Day coverage"]])) === upper(command.dayCoverage);
        const startMatch = command.dayCoverage === "Partial day"
          ? timeFromCell(row[index["Start time"]]) === requireTime(cleanText(command.startTime), "Start time")
          : true;
        const endMatch = command.dayCoverage === "Partial day"
          ? timeFromCell(row[index["End time"]]) === requireTime(cleanText(command.endTime), "End time")
          : true;
        if (coverageMatch && startMatch && endMatch) matches.push(position);
      } else if (upper(String(row[index["Away type"]])) === upper(command.awayType)) {
        matches.push(position);
      }
    });

    if (matches.length > 1) {
      throw new Error(`Ambiguous workbook identity: ${matches.length} rows match the request.`);
    }
    if (matches.length === 0 && command.operation === "CANCEL") {
      throw new Error("The approved workbook row to cancel was not found.");
    }

    const workbookStatus = command.requestType === "PTO"
      ? (command.operation === "CANCEL" ? "Cancelled" : "Approved")
      : deriveAwayStatus(command);
    if (matches.length === 1) {
      const rowPosition = matches[0];
      const current = rows[rowPosition];
      const replacement = command.requestType === "PTO"
        ? buildPtoRow(headers, index, command, current)
        : buildAwayRow(headers, index, command, current);
      const unchanged = replacement.every((value, position) => String(value) === String(current[position]));
      const rowRange = table.getRangeBetweenHeaderAndTotal().getRow(rowPosition);
      if (!unchanged) rowRange.setValues([replacement]);
      applyFormats(rowRange, index, command.requestType);
      return {
        ok: true, requestId, action: unchanged ? "NO_CHANGE" : "UPDATED",
        tableName, workbookStatus, matchCount: 1,
        message: unchanged ? "The exact workbook row was already current." : "The exact workbook row was updated."
      };
    }

    const newRow = command.requestType === "PTO"
      ? buildPtoRow(headers, index, command)
      : buildAwayRow(headers, index, command);
    table.addRow(-1, newRow);
    const addedRange = table.getRangeBetweenHeaderAndTotal().getLastRow();
    applyFormats(addedRange, index, command.requestType);
    return {
      ok: true, requestId, action: "INSERTED", tableName, workbookStatus,
      matchCount: 0, message: "A new approved workbook row was inserted."
    };
  } catch (error) {
    return {
      ok: false, requestId, action: "ERROR", tableName, workbookStatus: "",
      matchCount: 0, message: error instanceof Error ? error.message : String(error)
    };
  }
}
