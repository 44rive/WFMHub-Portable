declare namespace ExcelScript {
  interface Workbook {
    getTables(): Table[];
  }

  interface Table {
    getName(): string;
    getHeaderRowRange(): Range;
    getRowCount(): number;
    getRangeBetweenHeaderAndTotal(): Range;
    addRow(index: number, values: (string | number | boolean)[]): void;
  }

  interface Range {
    getTexts(): string[][];
    getValues(): (string | number | boolean)[][];
    getRow(index: number): Range;
    getLastRow(): Range;
    getCell(row: number, column: number): Range;
    setValues(values: (string | number | boolean)[][]): void;
    setNumberFormat(numberFormat: string[][]): void;
  }
}
