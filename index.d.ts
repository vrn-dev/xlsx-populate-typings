export class XlsxPopulate {
  /**
   * The XLSX mime type.
   * @type {string}
   */
  MIME_TYPE: string;

  /**
   * The XLSX error class.
   * @type {FormulaError}
   */
  public FormulaError: FormulaError;

  /**
   * Convert a date to a number for Excel.
   * @param {Date} date - The date.
   * @returns {number} The number.
   */
  static dateToNumber(date: Date): number;

  /**
   * Convert an Excel number to a date.
   * @param {number} num - The number.
   * @returns {Date} The date.
   */
  static numberToDate(num: number): Date;

  /**
   * Create a new blank workbook.
   * @returns {Promise<Workbook>} The workbook.
   */
  static fromBlankAsync(): Promise<Workbook>;

  /**
   * Loads a workbook from a data object. (Supports any supported [JSZip data types]{@link https://stuk.github.io/jszip/documentation/api_jszip/load_async.html}).
   * @param {string | number[] | ArrayBuffer | Uint8Array | Buffer | Blob | Promise<any>} data - The data to load.
   * @param {{password: string}} [opts] - Options
   * {string} [password] - The password to decrypt the workbook.
   * @returns {Promise<Workbook>} The workbook.
   */
  static fromDataAsync(data: string | number[] | ArrayBuffer | Uint8Array | Buffer | Blob | Promise<any>, opts?: { password: string }): Promise<Workbook>;

  /**
   * Loads a workbook from file.
   * @param {string} path - The path to the workbook.
   * @param {{password: string}} [opts] - Options
   * {string} [password] - The password to decrypt the workbook.
   * @returns {Promise<Workbook>} The workbook.
   */
  static fromFileAsync(path: string, opts?: { password: string }): Promise<Workbook>;
}

export class Workbook {
  static MIME_TYPE: string;

  /**
   * Create a new blank Workbook.
   * @returns {Promise<Workbook>} The workbook.
   */
  static fromBlankAsync(): Promise<Workbook>;

  /**
   * Loads a workbook from a data object. (Supports any supported [JSZip data types]{@link https://stuk.github.io/jszip/documentation/api_jszip/load_async.html}).
   * @param {string | number[] | ArrayBuffer | Uint8Array | Buffer | Blob | Promise<any>} data - The data to load.
   * @param {{password: string}} [opts] - Options: The password to decrypt the workbook.
   * @returns {Promise<Workbook>} The workbook.
   */
  static fromDataAsync(data: string | number[] | ArrayBuffer | Uint8Array | Buffer | Blob | Promise<any>, opts?: { password: string }): Promise<Workbook>;

  /**
   * Loads a workbook from file.
   * @param {string} path - The path to the workbook.
   * @param {{password: string}} [opts] - Options: The password to decrypt the workbook.
   * @returns {Promise<Workbook>} The workbook.
   */
  static fromFileAsync(path: string, opts?: { password: string }): Promise<Workbook>;

  /**
   * Get the active sheet in the workbook.
   * @returns {Sheet} The active sheet.
   */
  activeSheet(): Sheet;

  /**
   * Set the active sheet in the workbook.
   * @param {Sheet | string | number} sheet - The sheet or name of sheet or index of sheet to activate. The sheet must not be hidden.
   * @returns {Workbook} The workbook.
   */
  activeSheet(sheet: Sheet | string | number): Workbook;

  /**
   * Add a new sheet to the workbook.
   * @param {string} name - The name of the sheet. Must be unique, less than 31 characters, and may not contain the following characters: \ / * [ ] : ?
   * @param {number | string | Sheet} [indexOrBeforeSheet] - The index to move the sheet to or the sheet (or name of sheet) to move this sheet before. Omit this argument to move to the end of the workbook.
   * @returns {Sheet} The new sheet.
   */
  addSheet(name: string, indexOrBeforeSheet?: number | string | Sheet): Sheet;

  /**
   * Gets a defined name scoped to the workbook.
   * @param {string} name - The defined name.
   * @returns {string | Cell | Range | Row | Column} What the defined name refers to or undefined if not found. Will return the string formula if not a Row, Column, Cell, or Range.
   */
  definedName(name: string): string | Cell | Range | Row | Column;

  /**
   * Set a defined name scoped to the workbook.
   * @param {string} name - The defined name.
   * @param {string | Cell | Range | Row | Column} refersTo - What the name refers to.
   * @returns {Workbook} The workbook.
   */
  definedName(name: string, refersTo: string | Cell | Range | Row | Column): Workbook;

  /**
   * Delete a sheet from the workbook.
   * @param {Sheet | string | number} sheet - The sheet or name of sheet or index of sheet to move.
   * @returns {Workbook} The workbook.
   */
  deleteSheet(sheet: Sheet | string | number): Workbook;

  /**
   * Find the given pattern in the workbook and optionally replace it.
   * @param {string | RegExp} pattern - The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches.
   * @param {string | Function} [replacement] - The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in each cell will be replaced.
   * @returns {boolean} A flag indicating if the pattern was found.
   */
  find(pattern: string | RegExp, replacement?: string | Function): boolean;

  /**
   * Move a sheet to a new position.
   * @param {Sheet | string | number} sheet - The sheet or name of sheet or index of sheet to move.
   * @param {number | string | Sheet} [indexOrBeforeSheet] - The index to move the sheet to or the sheet (or name of sheet) to move this sheet before. Omit this argument to move to the end of the workbook.
   * @returns {Workbook} The workbook.
   */
  moveSheet(sheet: Sheet | string | number, indexOrBeforeSheet?: number | string | Sheet): Workbook;

  /**
   * Generates the workbook output.
   * @param {string | {type?: string, password?: string}} [type] - The type of the data to return: base64, binarystring, uint8array, arraybuffer, blob, nodebuffer. Defaults to 'nodebuffer' in Node.js and 'blob' in browsers.
   * [opts] - Options
   * {string} [opts.type] - The type of the data to return: base64, binarystring, uint8array, arraybuffer, blob, nodebuffer. Defaults to 'nodebuffer' in Node.js and 'blob' in browsers.
   * {string} [opts.password] - The password to use to encrypt the workbook.
   * @returns {string | Uint8Array | ArrayBuffer | Blob | Buffer} The data.
   */
  outputAsync(type?: string | { type?: string; password?: string }): string | Uint8Array | ArrayBuffer | Blob | Buffer;

  /**
   * Gets the sheet with the provided name or index (0-based).
   * @param {string | number} sheetNameOrIndex - The sheet name or index.
   * @returns {Sheet} - The sheet or undefined if not found.
   */
  sheet(sheetNameOrIndex: string | number): Sheet;

  /**
   * Get an array of all the sheets in the workbook.
   * @returns {Sheet[]} The sheets.
   */
  sheets(): Sheet[];

  /**
   * Gets an individual property.
   * @param {string} name - The name of the property.
   * @returns {CoreProperties} The property.
   */
  /**
   * Gets multiple properties
   * @param {string[]} name - The names of the properties.
   * @returns {CoreProperties} Object whose keys are the property names and values are the properties.
   */
  property(name: string | string[]): CoreProperties;

  /**
   * Sets an individual property.
   * @param {string} name - The name of the property.
   * @param {*} value - The value to set.
   * @returns {Workbook} The workbook.
   */
  property(name: string, value: any): Workbook;

  /**
   * Sets multiple properties.
   * @param {CoreProperties} properties - Object whose keys are the property names and values are the values to set.
   * @returns {Workbook} The workbook.
   */
  property(properties: CoreProperties): Workbook;

  /**
   * Get access to core properties object
   * @returns {CoreProperties} The core properties.
   */
  properties(): CoreProperties;

  /**
   * Write the workbook to file. (Not supported in browsers).
   * @param {string} path - The path of the file to write.
   * @param {{password: string}} [opts] - The password to encrypt the workbook.
   * @returns {Promise<any>} A promise.
   */
  toFileAsync(path: string, opts?: { password: string }): Promise<any>;
}

interface KeyValObject {
  [key: string]: any;
}

interface CoreProperties {
  [key: string]: any;
}

export class Sheet {
  /**
   * Creates a new instance of Sheet.
   * @param {Workbook} workbook - The parent workbook.
   * @param idNode - THe sheet ID node (from the parent workbook).
   * @param node - The sheet node.
   * @param [relationshipsNode] - The optional sheet relationships node.
   */
  constructor(workbook: Workbook, idNode: any, node: any, relationshipsNode?: any);

  /**
   * Gets a value indicating whether the sheet is the active sheet in the workbook.
   * @returns {boolean} True if active, false otherwise.
   */
  active(): boolean;

  /**
   * Make the sheet the active sheet in the workbook.
   * @param {boolean} active - Must be set to `true`. Deactivating directly is not supported. To deactivate, you should activate a different sheet instead.
   * @returns {Sheet} The sheet.
   */
  active(active: boolean): Sheet;

  /**
   * Get the active cell in the sheet.
   * @returns {Cell | Sheet} The active cell.
   */
  activeCell(): Cell;

  /**
   * Set the active cell in the workbook.
   * @param {string | Cell} cell - The cell or address of cell to activate.
   * @returns {Sheet} The sheet.
   */
  activeCell(cell: string | Cell): Sheet;

  /**
   * Set the active cell in the workbook by row and column.
   * @param {number} rowNumber - The row number of the cell.
   * @param {string | number} columnNameOrNumber - The column name or number of the cell.
   * @returns {Sheet} The sheet.
   */
  activeCell(rowNumber: number, columnNameOrNumber: string | number): Sheet;

  /**
   * Gets the cell with the given address.
   * @param {string} address - The address of the cell.
   * @returns {Cell} The cell.
   */
  cell(address: string): Cell;

  /**
   * Gets the cell with the given row and column numbers.
   * @param {number} rowNumber - The row number of the cell.
   * @param {string | number} columnNameOrNumber - The column name or number of the cell.
   * @returns {Cell}
   */
  cell(rowNumber: number, columnNameOrNumber: string | number): Cell;

  /**
   * Gets a column in the sheet.
   * @param {string | number} columnNameOrNumber - The name or number of the column.
   * @returns {Column} The column.
   */
  column(columnNameOrNumber: string | number): Column;

  /**
   * Gets a defined name scoped to the sheet.
   * @param {string} name - The defined name.
   * @returns {string | Cell | Range | Row | Column} What the defined name refers to or undefined if not found. Will return the string formula if not a Row, Column, Cell, or Range.
   */
  definedName(name: string): string | Cell | Range | Row | Column;

  /**
   * Set a defined name scoped to the sheet.
   * @param {string} name - The defined name.
   * @param {string | Cell | Range | Row | Column} refersTo - What the name refers to.
   * @returns {Workbook} The workbook.
   */
  definedName(name: string, refersTo: string | Cell | Range | Row | Column): Workbook;

  /**
   * Deletes the sheet and returns the parent workbook.
   * @returns {Workbook} The workbook.
   */
  delete(): Workbook;

  /**
   * Find the given pattern in the sheet and optionally replace it.
   * @param {string | RegExp} pattern - The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches.
   * @param {string | Function} [replacement] - The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in each cell will be replaced.
   * @returns {Cell[]} The matching cells.
   */
  find(pattern: string | RegExp, replacement?: string | Function): Cell[];

  /**
   * Gets a value indicating whether this sheet's grid lines are visible.
   * @returns {boolean} True if selected, false if not.
   */
  gridLinesVisible(): boolean;

  /**
   * Sets whether this sheet's grid lines are visible.
   * @param {boolean} selected - True to make visible, false to hide.
   * @returns {Sheet} The sheet.
   */
  gridLinesVisible(selected: boolean): Sheet;

  /**
   * Gets a value indicating if the sheet is hidden or not.
   * @returns {boolean|string} True if hidden, false if visible, and 'very' if very hidden.
   */
  hidden(): boolean | string;

  /**
   * Set whether the sheet is hidden or not.
   * @param {boolean | string} hidden - True to hide, false to show, and 'very' to make very hidden.
   * @returns {Sheet} The sheet.
   */
  hidden(hidden: boolean | string): Sheet;

  /**
   * Move the sheet.
   * @param {number | string | Sheet} [indexOrBeforeSheet] - The index to move the sheet to or the sheet (or name of sheet) to move this sheet before. Omit this argument to move to the end of the workbook.
   * @returns {Sheet} The sheet.
   */
  move(indexOrBeforeSheet?: number | string | Sheet): Sheet;

  /**
   * Get the name of the sheet.
   * @returns {string} The sheet name.
   */
  name(): string;

  /**
   * Set the name of the sheet. *Note: this method does not rename references to the sheet so formulas, etc. can be broken. Use with caution!*
   * @param {string} name - The name to set to the sheet.
   * @returns {Sheet} The sheet.
   */
  name(name: string): Sheet;

  /**
   * Gets a range from the given range address.
   * @param {string} address - The range address (e.g. 'A1:B3').
   * @returns {Range} The range.
   */
  range(address: string): Range;

  /**
   * Gets a range from the given cells or cell addresses.
   * @param {string | Cell} startCell - The starting cell or cell address (e.g. 'A1').
   * @param {string | Cell} endCell - The ending cell or cell address (e.g. 'B3').
   * @returns {Range} The range.
   */
  range(startCell: string | Cell, endCell: string | Cell): Range;

  /**
   * Gets a range from the given row numbers and column names or numbers.
   * @param {number} startRowNumber - The starting cell row number.
   * @param {string | number} startColumnNameOrNumber - The starting cell column name or number.
   * @param {number} endRowNumber - The ending cell row number.
   * @param {string | number} endColumnNameOrNumber - The ending cell column name or number.
   * @returns {Range} The range.
   */
  range(startRowNumber: number, startColumnNameOrNumber: string | number, endRowNumber: number, endColumnNameOrNumber: string | number): Range;

  /**
   * Sets or Unsets sheet autoFilter to a Range.
   * @param {Range} [range] - The autoFilter range. Omit param to unset.
   * @returns {Sheet} The sheet.
   */
  autoFilter(range?: Range): Sheet;

  /**
   * Gets the row with the given number.
   * @param {number} rowNumber - The row number.
   * @returns {Row} The row with the given number.
   */
  row(rowNumber: number): Row;

  /**
   * Get the tab color. (See style [Color](#color).)
   * @returns {Color} The color or undefined if not set.
   */
  tabColor(): Color;

  /**
   * Sets or Deletes the tab color. (See style [Color](#color).)
   * @param {Color | string | number | null} color - Color of the tab. If string, will set an RGB color. If number, will set a theme color. If Null will delete the tabColor.
   * @returns {Color | string | number} The color.
   */
  tabColor(color: Color | string | number | null): Color | string | number;

  /**
   * Gets a value indicating whether this sheet is selected.
   * @returns {boolean} True if selected, false if not.
   */
  tabSelected(): boolean;

  /**
   * Sets whether this sheet is selected.
   * @param {boolean} selected - True to select, false to deselected.
   * @returns {Sheet} The sheet.
   */
  tabSelected(selected: boolean): Sheet;

  /**
   * Get the range of cells in the sheet that have contained a value or style at any point. Useful for extracting the entire sheet contents.
   * @returns {Range | undefined} - The used range or undefined if no cells in the sheet are used.
   */
  usedRange(): Range | undefined;

  /**
   * Gets the parent workbook.
   * @returns {Workbook} The parent workbook.
   */
  workbook(): Workbook;
}

export class Range {
  /**
   * Creates a new instance of Range.
   * @param {Cell} startCell - The start cell.
   * @param {Cell} endCell - The end cell.
   */
  constructor(startCell: Cell, endCell: Cell);

  /**
   * Get the address of the range.
   * @param {{includeSheetName: boolean, startRowAnchored: boolean, startColumnAnchored: boolean, endRowAnchored: boolean, endColumnAnchored: boolean, anchored: boolean}} [opts] - Options
   * Options:
   * {boolean} [includeSheetName] - Include the sheet name in the address.
   * {boolean} [startRowAnchored] - Anchor the start row.
   * {boolean} [startColumnAnchored] - Anchor the start column.
   * {boolean} [endRowAnchored] - Anchor the end row.
   * {boolean} [endColumnAnchored] - Anchor the end column.
   * {boolean} [anchored] - Anchor all row and columns.
   * @returns {string} The address
   */
  address(opts?: { includeSheetName: boolean, startRowAnchored: boolean, startColumnAnchored: boolean, endRowAnchored: boolean, endColumnAnchored: boolean, anchored: boolean }): string;

  /**
   * Gets a cell within the range.
   * @param {number} rowIndex - Row index relative to the top-left corner of the range (0-based).
   * @param {number} columnIndex - Column index relative to the top-left corner of the range (0-based).
   * @returns {Cell} The cell.
   */
  cell(rowIndex: number, columnIndex: number): Cell;

  /**
   * Sets sheet autoFilter to this range.
   * @returns {Range} This range.
   */
  autoFilter(): Range;

  /**
   * Get the cells in the range as a 2D array.
   * @returns {Array<Cell[]>} The cells.
   */
  cells(): Array<Cell[]>;

  /**
   * Clear the contents of all the cells in the range.
   * @returns {Range} The range.
   */
  clear(): Range;

  /**
   * Get the end cell of the range.
   * @returns {Cell} The end cell.
   */
  endCell(): Cell;

  /**
   * Call a function for each cell in the range. Goes by row then column.
   * @param {(cell: Cell, rowIndex: number, columnIndex: number) => Range} callback: Function called for each cell in the range.
   * @returns {Range} The range.
   */
  forEach(callback: (cell: Cell, rowIndex: number, columnIndex: number) => Range): Range;

  /**
   * Gets the shared formula in the start cell (assuming it's the source of the shared formula).
   * @returns {string}
   */
  formula(): string;

  /**
   * Sets the shared formula in the range. The formula will be translated for each cell.
   * @param {string} formula - The formula to set.
   * @returns {Range} The range.
   */
  formula(formula: string): Range;

  /**
   * Creates a 2D array of values by running each cell through a callback.
   * @param {(cell: Cell, rowIndex: number, columnIndex: number, range: Range) => any} callback - Function called for each cell in the range.
   * @returns {Array<any[]>} The 2D array of return values.
   */
  map(callback: (cell: Cell, rowIndex: number, columnIndex: number, range: Range) => any): Array<any[]>;

  /**
   * Gets a value indicating whether the cells in the range are merged.
   * @returns {boolean} The value
   */
  merged(): boolean;

  /**
   * Sets a value indicating whether the cells in the range should be merged.
   * @param {boolean} merged - True to merge, false to unmerge.
   * @returns {Range} The range.
   */
  merged(merged: boolean): Range;

  /**
   * Gets the data validation object attached to the Range.
   * @returns {object}
   */
  dataValidation(): object;

  /**
   * Set or clear the data validation object of the entire range.
   * @param {object | null} dataValidation - Object or null to clear.
   * @returns {Range}
   */
  dataValidation(dataValidation: object | null): Range;

  /**
   * Reduces the range to a single value accumulated from the result of a function called for each cell.
   * @param {(accumulator: any, cell: Cell, rowIndex: number, columnIndex: number, range: Range) => any} callback - Function called for each cell in the range.
   * @param initialValue - The initial value.
   * @returns {any} The accumulated value.
   */
  reduce(callback: (accumulator: any, cell: Cell, rowIndex: number, columnIndex: number, range: Range) => any, initialValue?: any): any;

  /**
   * Gets the parent sheet of the range.
   * @returns {Sheet} The parent sheet.
   */
  sheet(): Sheet;

  /**
   * Gets the start cell of the range.
   * @returns {Cell} The start cell.
   */
  startCell(): Cell;

  /**
   * Gets a single style for each cell.
   * @param {string} name - The name of the style.
   * @returns {Array<any[]>} 2D array of style values.
   */
  style(name: string): Array<any[]>;

  /**
   * Gets multiple styles for each cell.
   * @param {string[]} names - The names of the styles.
   * @returns {Styles} Object whose keys are style names and values are 2D arrays of style values.
   */
  style(names: string[]): Styles;

  /**
   * Set the style in each cell to the result of a function called for each.
   * @param {string} name - The name of the style.
   * @param {(cell: Cell, rowIndex: number, columnIndex: number, range: Range) => any} values - The callback to provide value for the cell.
   * @returns {Range} The range.
   */
  style(name: string, values: (cell: Cell, rowIndex: number, columnIndex: number, range: Range) => any): Range;

  // noinspection TsLint
  /**
   * Sets the style in each cell to the corresponding value in the given 2D array of values.
   * @param {string} name - The name of the style.
   * @param {Array<any[]>} values - The style values to set.
   * @returns {Range} The range.
   */
  style(name: string, values: Array<any[]>): Range;

  // noinspection TsLint
  /**
   * Set the style of all cells in the range to a single style value.
   * @param {string} name - The name of the style.
   * @param value - The style value to set.
   * @returns {Range} The range.
   */
  style(name: string, value: any): Range;

  /**
   * Set multiple styles for the cells in the range.
   * @param {Styles} styles - Style object.
   * @returns {Range} The range.
   */
  style(styles: Styles): Range;

  /**
   * Invoke a callback on the range and return the range. Useful for method chaining.
   * @param {(range: Range) => void} callback - The callback function.
   * @returns {Range} The range.
   */
  tap(callback: (range: Range) => void): Range;

  /**
   * Invoke a callback on the range and return the value provided by the callback. Useful for method chaining.
   * @param {(range: Range) => any} callback - The callback function.
   * @returns {any} The return value of the callback.
   */
  thru(callback: (range: Range) => any): any;

  /**
   * Get the values of each cell in the range as a 2D array.
   * @returns {Array<any[]>} The values
   */
  value(): Array<any[]>;

  /**
   * Set the values in each cell to the result of a function called for each.
   * Sets the value in each cell to the corresponding value in the given 2D array of values.
   * Set the value of all cells in the range to a single value.
   * @param {(cell: Cell, rowIndex: number, columnIndex: number, range: Range) => (any | Array<any[]>)} values - The callback to provide value for the cell / The value/s to set.
   * @returns {Range} The range.
   */
  value(values: (cell: Cell, rowIndex: number, columnIndex: number, range: Range) => any | Array<any[]> | any): Range;

  /**
   * Gets the parent workbook
   * @returns {Workbook}
   */
  workbook(): Workbook;
}

export class Cell {
  /**
   * Creates a new instance of cell.
   * @param {Row} row - The parent row.
   * @param node - The cell node.
   * @param [styleId] - The style ID for the new cell
   */
  constructor(row: Row, node: any, styleId?: any);

  /**
   * Gets a value indicating whether the cell is the active cell in the sheet.
   * @returns {boolean} True if active, false otherwise.
   */
  active(): boolean;

  /**
   * Make the cell the active cell in the sheet.
   * @param {boolean} active - Must be set to `true`. Deactivating directly is not supported. To deactivate, you should activate a different cell instead.
   * @returns {Cell} The cell.
   */
  active(active: boolean): Cell;

  /**
   * Get the address of the column.
   * @param {{includeSheetName: boolean, rowAnchored: boolean, columnAnchored: boolean, anchored: boolean}} [opts] - Options
   * {boolean} [includeSheetName] - Include the sheet name in the address.
   * {boolean} [rowAnchored] - Anchor the row.
   * {boolean} [columnAnchored] - Anchor the column.
   * {boolean} [anchored] - Anchor both the row and the column.
   * @returns {string} The address
   */
  address(opts?: { includeSheetName: boolean; rowAnchored: boolean; columnAnchored: boolean; anchored: boolean; }): string;

  /**
   * Gets the parent column of the cell.
   * @returns {Column} The parent column
   */
  column(): Column;

  /**
   * Clears the contents from the cell.
   * @returns {Cell} The cell.
   */
  clear(): Cell;

  /**
   * Gets the column name of the cell.
   * @returns {string} The column name.
   */
  columnName(): string;

  /**
   * Gets the column number of the cell (1-based).
   * @returns {number} The column number.
   */
  columnNumber(): number;

  /**
   * Find the given pattern in the cell and optionally replace it.
   * @param {string | RegExp} pattern - The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches.
   * @param {string | Function} [replacement] - The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in the cell will be replaced.
   * @returns {boolean} A flag indicating if the pattern was found.
   */
  find(pattern: string | RegExp, replacement: string | Function): boolean;

  /**
   * Gets the formula in the cell. Note that if a formula was set as part of a range, the getter will return 'SHARED'. This is a limitation that may be addressed in a future release.
   * @returns {string} The formula in the cell.
   */
  formula(): string;

  /**
   * Sets the formula in the cell.
   * @param {string} formula - The formula to set.
   * @returns {Cell} The cell.
   */
  formula(formula: string): Cell;

  /**
   * Gets the hyperlink attached to the cell.
   * @returns {string} The hyperlink or undefined if not set.
   */
  hyperlink(): string;

  /**
   * Set or clear the hyperlink on the cell.
   * @param {string} hyperlink - The hyperlink to set or null to clear.
   * @returns {Cell}
   */
  hyperlink(hyperlink: string | null): Cell;

  /**
   * Gets the data validation object attached to the cell.
   * @returns {object} The data validation or undefined if not set.
   */
  dataValidation(): object;

  /**
   * Set or clear the data validation object of the cell.
   * @param {object | null} dataValidation - Object or null to clear.
   * @returns {Cell} The cell.
   */
  dataValidation(dataValidation: object | null): Cell;

  /**
   * Invoke a callback on the cell and return the cell. Useful for method chaining.
   * @param {(cell: Cell) => void} callback - The callback function.
   * @returns {Cell} The cell.
   */
  tap(callback: (cell: Cell) => void): Cell;

  /**
   * Invoke a callback on the cell and return the value provided by the callback. Useful for method chaining.
   * @param {(cell: Cell) => any} callback - The callback function.
   * @returns {any} The return value of the callback.
   */
  thru(callback: (cell: Cell) => any): any;

  /**
   * Create a range from this cell and another.
   * @param {Cell | string} cell - The other cell or cell address to range to.
   * @returns {Range} The range.
   */
  rangeTo(cell: Cell | string): Range;

  /**
   * Returns a cell with a relative position given the offsets provided.
   * @param {number} rowOffset - The row offset (0 for the current row).
   * @param {number} columnOffset - The column offset (0 for the current column).
   * @returns {Cell} The relative cell.
   */
  relativeCell(rowOffset: number, columnOffset: number): Cell;

  /**
   * Gets the parent row of the cell.
   * @returns {Row} The parent row.
   */
  row(): Row;

  /**
   * Gets the row number of the cell (1-based).
   * @returns {number} The row number.
   */
  rowNumber(): number;

  /**
   * Gets the parent sheet.
   * @returns {Sheet} The parent sheet.
   */
  sheet(): Sheet;

  /**
   * Gets an individual style.
   * @param {string} name - The name of the style.
   * @returns {any} The style
   */
  style(name: string): any;

  /**
   * Gets multiple styles.
   * @param {string[]} names - Gets multiple styles.
   * @returns {{[p: string]: any}} Object whose keys are the style names and values are the styles.
   */
  style(names: string[]): { [key: string]: any };

  /**
   * Sets an individual style.
   * @param {string} name - The name of the style.
   * @param value - The value to set.
   * @returns {Cell} The cell.
   */
  style(name: string, value: any): Cell;

  /**
   * Sets the styles in the range starting with the cell.
   * @param {string} name -  The name of the style.
   * @param {Array<any[]>} values - 2D array of values to set.
   * @returns {Range} The range that was set.
   */
  style(name: string, values: Array<any[]>): Range;

  /**
   * Sets multiple styles / Sets to a specific style
   * @param {Styles | Style} styles - Object whose keys are the style names and values are the styles to set / Style object given from stylesheet.createStyle.
   * @returns {Cell} The cell.
   */
  style(styles: Styles | Style): Cell;

  /**
   * Gets the value of the cell
   * @returns {string | boolean | number | Date}
   */
  value(): string | boolean | number | Date;

  /**
   * Sets the value of the cell
   * @param {string | boolean | number} value
   * @returns {Cell}
   */
  value(value: string | boolean | number): Cell;

  /**
   * Sets the values in the range starting with the cell
   * @param {Array<string[] | boolean[] | number[]>} values
   * @returns {Range}
   */
  value(values: Array<string[] | boolean[] | number[]>): Range;

  /**
   * Gets the parent workbook.
   * @returns {Workbook} The parent workbook.
   */
  workbook(): Workbook;
}

export class Row {
  /**
   * Creates a new instance of Row.
   * @param {Sheet} sheet - The parent sheet.
   * @param node - The row node.
   */
  constructor(sheet: Sheet, node: any);

  /**
   * Get the address of the row.
   * @param {{includeSheetName: boolean; anchored: boolean}} [opts] - Options
   * {boolean} [includeSheetName] - Include the sheet name in the address.
   * {boolean} [anchored] - Anchor the address.
   * @returns {string} The address.
   */
  address(opts?: { includeSheetName: boolean, anchored: boolean }): string;

  /**
   * Get a cell in the row.
   * @param {string | number} columnNameOrNumber - The name or number of the column.
   * @returns {Cell} The cell.
   */
  cell(columnNameOrNumber: string | number): Cell;

  /**
   * Gets the row height.
   * @returns {number} The height (or undefined).
   */
  height(): number;

  /**
   * Sets or clears the row height.
   * @param {number | null} height - The height of the row or null to clear it.
   * @returns {Row} The row.
   */
  height(height: number | null): Row;

  /**
   * Gets a value indicating whether the row is hidden.
   * @returns {boolean} A flag indicating whether the row is hidden.
   */
  hidden(): boolean;

  /**
   * Sets whether the row is hidden.
   * @param {boolean} hidden - A flag indicating whether to hide the row.
   * @returns {Row}
   */
  hidden(hidden: boolean): Row;

  /**
   * Gets the row number.
   * @returns {number} The row number.
   */
  rowNumber(): number;

  /**
   * Gets the parent sheet of the row.
   * @returns {Sheet} The parent sheet.
   */
  sheet(): Sheet;

  /**
   * Gets an individual style.
   * @param {string} name - The name of the style.
   * @returns {any} The style.
   */
  style(name: string): any;

  /**
   * Gets multiple styles.
   * @param {string[]} names - The names of the style.
   * @returns {{[p: string]: any}} Object whose keys are the style names and values are the styles.
   */
  style(names: string[]): { [key: string]: any };

  /**
   * Sets an individual style.
   * @param {string} name - The name of the style.
   * @param value - The value to set.
   * @returns {Cell} The cell.
   */
  style(name: string, value: any): Cell;

  /**
   * Sets multiple styles / Sets to a specific style
   * @param {Styles | Style} styles - Object whose keys are the style names and values are the styles to set / Style object given from stylesheet.createStyle.
   * @returns {Cell} The cell.
   */
  style(styles: Styles | Style): Cell;

  /**
   * Get the parent workbook.
   * @returns {Workbook} The parent workbook.
   */
  workbook(): Workbook;
}

export class Column {
  /**
   * Creates a new Column.
   * @param {Sheet} sheet - The parent sheet.
   * @param node - The column node.
   * @private
   */
  constructor(sheet: Sheet, node: any);

  /**
   * Get the address of the column.
   * @param {{includeSheetName: boolean; anchored: boolean}} [opts] - Options
   * {boolean} [includeSheetName] - Include the sheet name in the address.
   * {boolean} [anchored] - Anchor the address.
   * @returns {string} The address
   */
  address(opts?: { includeSheetName: boolean; anchored: boolean; }): string;

  /**
   * Get a cell within the column.
   * @param {number} rowNumber - The row number.
   * @returns {Cell} The cell in the column with the given row number.
   */
  cell(rowNumber: number): Cell;

  /**
   * Get the name of the column.
   * @returns {string} The column name.
   */
  columnName(): string;

  /**
   * Get the number of the column.
   * @returns {number} The column number.
   */
  columnNumber(): number;

  /**
   * Gets a value indicating whether the column is hidden.
   * @returns {boolean}  A flag indicating whether the column is hidden.
   */
  hidden(): boolean;

  /**
   * Sets whether the column is hidden.
   * @param {boolean} hidden - A flag indicating whether to hide the column.
   * @returns {Column} The column
   */
  hidden(hidden: boolean): Column;

  /**
   * Get the parent sheet.
   * @returns {Sheet} The parent sheet.
   */
  sheet(): Sheet;

  /**
   * Gets an individual style.
   * @param {string} name - The name of the style.
   * @returns {any} The style.
   */
  style(name: string): any;

  /**
   * Gets multiple styles.
   * @param {string[]} names - The names of the style.
   * @returns {{[p: string]: any}} Object whose keys are the style names and values are the styles.
   */
  style(names: string[]): { [key: string]: any };

  /**
   * Sets an individual style.
   * @param {string} name - The name of the style.
   * @param value - The value to set.
   * @returns {Cell} The cell.
   */
  style(name: string, value: any): Cell;

  /**
   * Sets multiple styles / Sets to a specific style
   * @param {Styles | Style} styles - Object whose keys are the style names and values are the styles to set / Style object given from stylesheet.createStyle.
   * @returns {Cell} The cell.
   */
  style(styles: Styles | Style): Cell;

  /**
   * Gets the width.
   * @returns {number} The width (or undefined).
   */
  width(): number;

  /**
   * Sets the width.
   * @param {number} width - The width of the column.
   * @returns {Column} The column.
   */
  width(width: number): Column;

  /**
   * Get the parent workbook.
   * @returns {Workbook} The parent workbook.
   */
  workbook(): Workbook;
}

interface Color {
  rgb?: any;
  indexed?: any;
  theme?: any;
  tint?: any;
}

export class Style {
  /**
   * Creates a new instance of _Style.
   * @param {StyleSheet} styleSheet - The stylesheet.
   * @param {number} id - The style ID.
   * @param xfNode - The xf node.
   * @param fontNode - The font node.
   * @param fillNode - The fill node.
   * @param borderNode - the border node.
   */
  constructor(styleSheet: StyleSheet, id: number, xfNode: any, fontNode: any, fillNode: any, borderNode: any);

  /**
   * Gets the style ID.
   * @returns {number} The ID.
   */
  id(): number;

  /**
   * Get a style.
   * @param {string} name - The style name.
   * @returns {any} The value of the style
   */
  style(name: string): any;

  /**
   * Set a style.
   * @param {string} name - The style name.
   * @param value - the value to set.
   * @returns {Style} The style
   */
  style(name: string, value: any): Style;
}

export class StyleSheet {
  /**
   * Creates an instance of _StyleSheet.
   * @param {string} node - The style sheet node
   */
  constructor(node: string);

  /**
   * Create a style.
   * @param {number} [sourceId] - The source style ID to copy, if provided.
   * @returns {Style} The style.
   */
  createStyle(sourceId?: number): Style;

  /**
   * Get the number format code for a given ID.
   * @param {number} id - The number format ID.
   * @returns {string} The format code.
   */
  getNumberFormatCode(id: number): string;

  /**
   * Get the number format ID for a given code.
   * @param {string} code - The format code.
   * @returns {number} The number format ID.
   */
  getNumberFormatId(code: string): number;
}

export class FormulaError {
  /**
   * Creates a new instance of Formula Error
   * @param {string} error - The error code.
   */
  constructor(error: string);

  /**
   * Get the error code.
   * @returns {string} The error code.
   */
  error(): string;
}

// noinspection TsLint
export interface Styles {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  subscript?: boolean;
  superscript?: boolean;
  fontSize?: number;
  fontFamily?: string;
  fontColor?: string;
  horizontalAlignment?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  justifyLastLine?: boolean;
  indent?: number;
  verticalAlignment?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
  wrapText?: boolean;
  shrinkToFit?: boolean;
  textDirection?: 'left-to-right' | 'right-to-left';
  textRotation?: number;
  angleTextCounterclockwise?: boolean;
  angleTextClockwise?: boolean;
  rotateTextUp?: boolean;
  rotateTextDown?: boolean;
  verticalText?: boolean;
  fill?: string;
  border?: string;
  borderColor?: string;
  borderStyle?: 'hair' | 'dotted' | 'dashDotDot' | 'dashed' | 'mediumDashDotDot' | 'thin' | 'slantDashDot' | 'mediumDashDot' | 'mediumDashed' | 'medium' | 'thick' | 'double';
  leftBorder?: boolean;
  rightBorder?: boolean;
  topBorder?: boolean;
  bottomBorder?: boolean;
  diagonalBorder?: boolean;
  leftBorderColor?: string;
  rightBorderColor?: string;
  topBorderColor?: string;
  bottomBorderColor?: string;
  diagonalBorderColor?: string;
  leftBorderStyle?: string;
  rightBorderStyle?: string;
  topBorderStyle?: string;
  bottomBorderStyle?: string;
  diagonalBorderStyle?: string;
  diagonalBorderDirection?: 'up' | 'down' | 'both';
  numberFormat?: '"AED "#,##0.00' | 'General' | '0' | '0.00' | '# |##0' | '# |##0.00' | '0%' | '0.00%' | '0.00E+00' | '# ?/?' | '# ??/??' | 'mm-dd-yy' | 'd-mmm-yy' | 'd-mmm' | 'mmm-yy' | 'h:mm AM/PM' | 'h:mm:ss AM/PM' | 'h:mm' | 'h:mm:ss' | 'm/d/yy h:mm' | '# |##0 ;(# |##0)' | '# |##0 ;[Red](# |##0)' | '# |##0.00;(# |##0.00)' | '# |##0.00;[Red](# |##0.00)' | 'mm:ss' | '[h]:mm:ss' | 'mmss.0' | '##0.0E+0' | '@';
}
