/**
 * Represents a unique coordinate on the spreadsheet.
 */
export interface SpreadsheetCoordinate {
  x: number;
  y: number;
  worksheet: string;
}

export interface CellInfo extends SpreadsheetCoordinate {
  orig: string;
  err: string;
}

/**
 * Represents spreadsheet *outputs*.
 */
export interface OutputInfo extends CellInfo {
  formula: string;
}

/**
 * Represents spreadsheet *inputs*.
 */
export interface InputInfo extends CellInfo {
  outputs: {
    x: number;
    y: number;
    worksheet: string;
    noerr: string;
  }[];
  style: CellStyle;
}

export interface CellStyle {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  'font-face': string;
  'font-size': number;
}

/**
 * Represents one question's information.
 */
export interface QuestionInfo {
  errors: InputInfo[];
  outputs: OutputInfo[];
}
