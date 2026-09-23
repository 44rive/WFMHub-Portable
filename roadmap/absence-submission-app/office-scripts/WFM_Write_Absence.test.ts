class MockRange implements ExcelScript.Range {
  constructor(
    private readonly table: MockTable,
    private readonly kind: "HEADER" | "BODY" | "ROW" | "CELL",
    private readonly rowIndex = -1,
    private readonly columnIndex = -1,
  ) {}

  getTexts(): string[][] {
    return this.getValues().map(row => row.map(value => String(value)));
  }

  getValues(): CellValue[][] {
    if (this.kind === "HEADER") return [this.table.headers.slice()];
    if (this.kind === "BODY") return this.table.rows.map(row => row.slice());
    if (this.kind === "ROW") return [this.table.rows[this.rowIndex].slice()];
    return [[this.table.rows[this.rowIndex][this.columnIndex]]];
  }

  getRow(index: number): ExcelScript.Range {
    return new MockRange(this.table, "ROW", index);
  }

  getLastRow(): ExcelScript.Range {
    return new MockRange(this.table, "ROW", this.table.rows.length - 1);
  }

  getCell(_row: number, column: number): ExcelScript.Range {
    return new MockRange(this.table, "CELL", this.rowIndex, column);
  }

  setValues(values: CellValue[][]): void {
    if (this.kind === "ROW") this.table.rows[this.rowIndex] = values[0].slice();
    if (this.kind === "CELL") this.table.rows[this.rowIndex][this.columnIndex] = values[0][0];
  }

  setNumberFormat(_numberFormat: string[][]): void {}
}

class MockTable implements ExcelScript.Table {
  public rows: CellValue[][] = [];

  constructor(public readonly name: string, public readonly headers: string[]) {}

  getName(): string { return this.name; }
  getHeaderRowRange(): ExcelScript.Range { return new MockRange(this, "HEADER"); }
  getRowCount(): number { return this.rows.length; }
  getRangeBetweenHeaderAndTotal(): ExcelScript.Range { return new MockRange(this, "BODY"); }
  addRow(_index: number, values: CellValue[]): void { this.rows.push(values.slice()); }
}

class MockWorkbook implements ExcelScript.Workbook {
  public readonly pto = new MockTable("tblFTEPTO", PTO_HEADERS);
  public readonly away = new MockTable("tblFTEAway", AWAY_HEADERS);
  getTables(): ExcelScript.Table[] { return [this.pto, this.away]; }
}

function assertEqual(actual: unknown, expected: unknown, label: string): void {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${String(expected)}, got ${String(actual)}`);
  }
}

const workbook = new MockWorkbook();
const fullDay = JSON.stringify({
  contractVersion: "1.0.0", requestId: "req-pto", operation: "UPSERT",
  requestType: "PTO", clientId: "00123", employeeName: "Jane Agent",
  startDate: "2026-09-15", endDate: "2026-09-16", dayCoverage: "Full day",
  startTime: "", endTime: "", awayType: "", comment: "Approved",
  today: "2026-09-10"
});

let result = main(workbook, fullDay);
assertEqual(result.ok, true, "PTO insert ok");
assertEqual(result.action, "INSERTED", "PTO inserted");
assertEqual(workbook.pto.rows.length, 1, "one PTO row");
assertEqual(workbook.pto.rows[0][0], "00123", "Client ID text preserved");
assertEqual(workbook.pto.rows[0][7], "PTO", "fixed PTO type");
assertEqual(workbook.pto.rows[0][8], "Approved", "approved status");

result = main(workbook, fullDay);
assertEqual(result.action, "NO_CHANGE", "PTO retry is idempotent");
assertEqual(workbook.pto.rows.length, 1, "retry creates no duplicate");

const cancelPto = JSON.stringify({...JSON.parse(fullDay), operation: "CANCEL"});
result = main(workbook, cancelPto);
assertEqual(result.action, "UPDATED", "PTO cancellation updates");
assertEqual(workbook.pto.rows[0][8], "Cancelled", "PTO cancelled status");

const openAway = JSON.stringify({
  contractVersion: "1.0.0", requestId: "req-away", operation: "UPSERT",
  requestType: "AWAY", clientId: "00123", employeeName: "Jane Agent",
  startDate: "2026-09-20", endDate: "", dayCoverage: "", startTime: "",
  endTime: "", awayType: "Long sickness", comment: "Open",
  today: "2026-09-10"
});
result = main(workbook, openAway);
assertEqual(result.action, "INSERTED", "Away inserted");
assertEqual(result.workbookStatus, "Active", "open Away is Active");
assertEqual(workbook.away.rows[0][5], "Active", "Away row Active");

const invalidPartial = JSON.stringify({
  contractVersion: "1.0.0", requestId: "req-invalid", operation: "UPSERT",
  requestType: "PTO", clientId: "00123", employeeName: "Jane Agent",
  startDate: "2026-09-15", endDate: "2026-09-15", dayCoverage: "Partial day",
  startTime: "10:00:00", endTime: "09:00:00", awayType: "", comment: "",
  today: "2026-09-10"
});
result = main(workbook, invalidPartial);
assertEqual(result.ok, false, "invalid partial PTO rejected");
assertEqual(workbook.pto.rows.length, 1, "invalid request writes nothing");

// The test runner prints one stable line for CI and manual checks.
declare const console: { log(message: string): void };
console.log("WFM_Write_Absence tests passed");
