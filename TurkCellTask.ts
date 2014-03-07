/**
 * Client-side logic for laying out a web page for a CheckCell error-ranking
 * question. Consumes a JSON blob and spits out a question in the desired div.
 */

// #region Dumb Browser Polyfills
// Who knows what browsers Mechanical Turk users are using?

// #endregion Dumb Browser Polyfills

// #region JSON Typings

/**
 * Represents a unique coordinate on the spreadsheet.
 */
interface SpreadsheetCoordinate {
  x: number;
  y: number;
  worksheet: string;
}

/**
 * Represents spreadsheet *outputs*.
 */
interface OutputInfo extends SpreadsheetCoordinate {
  orig: string;
  err: string;
}

/**
 * Represents spreadsheet *inputs*.
 */
interface InputInfo extends OutputInfo {
  outputs: {
    x: number;
    y: number;
    worksheet: string;
    noerr: string;
  }[];
}

/**
 * Represents one question's information.
 */
interface QuestionInfo {
  errors: InputInfo[];
  outputs: OutputInfo[];
}

/**
 * Data for a sample question.
 */
var sampleQuestion: QuestionInfo = {
  "errors": [
    {
      "x": 1,
      "y": 1,
      "worksheet": "sheet1",
      "orig": "12.343",
      "err": "123.43",
      "outputs": [
        {
          "x": 6,
          "y": 7,
          "worksheet": "sheet1",
          "noerr": "62.11"
        },
        {
          "x": 7,
          "y": 7,
          "worksheet": "sheet1",
          "noerr": "99.0"
        }
      ]
    },
    {
      "x": 1,
      "y": 2,
      "worksheet": "sheet1",
      "orig": "10.0",
      "err": "1.0",
      "outputs": [
        {
          "x": 6,
          "y": 7,
          "worksheet": "sheet1",
          "noerr": "0.0"
        },
        {
          "x": 7,
          "y": 7,
          "worksheet": "sheet1",
          "noerr": "99.0"
        }
      ]
    }
  ],
  "outputs": [
    {
      "x": 6,
      "y": 7,
      "worksheet": "sheet1",
      "orig": "0.0",
      "err": "100"
    },
    {
      "x": 7,
      "y": 7,
      "worksheet": "sheet1",
      "orig": "3.14159265359",
      "err": "99.0"
    }
  ]
};

// #endregion JSON Typings

// #region Helper Functions

/**
 * Converts a number into an Excel column ID.
 * e.g. 0 => A, 26 => AA.
 */
function getExcelColumn(i: number): string {
  var chars: string[] = [];
  do {
    chars.push(String.fromCharCode(65 + (i % 26)));
    i = Math.floor(i / 26) - 1;
  } while (i > -1);
  chars.reverse();
  return chars.join('');
}

/**
 * Tests an assertion.
 */
function assert(test: boolean, errorMessage: string = "") {
  if (!test) {
    throw new Error('Assertion error: ' + errorMessage);
  }
}

// #endregion Helper Functions

/**
 * Defines the different types of DD items.
 */
enum DDType { OUTPUT, INPUT }

/**
 * An item in the data dependency graph.
 */
interface DDItem {
  /**
   * Used for casting. Get the type of the item (output/input).
   */
  getType(): DDType;
  /**
   * Add an event listener to this item.
   * NOTE: We only support 'change' events.
   */
  addEventListener(type: string, cb: (data: string) => void): void;
  /**
   * Get the current value of the item.
   */
  getValue(): string;
}

class ChangeObservable {
  private listeners: { (data: string): void }[] = [];
  public addEventListener(type: string, cb: (data: string) => void): void {
    if (type === 'changed') {
      this.listeners.push(cb);
    }
  }

  public fireEvent(data: string) {
    var i: number;
    for (i = 0; i < this.listeners.length; i++) {
      this.listeners[i](data);
    }
  }
}

enum OutputStatus { ORIG, ERR, CUSTOM }

/**
 * Object instantiation of an OutputInfo from the JSON structure.
 */
class OutputItem extends ChangeObservable implements DDItem {
  private orig: string;
  private err: string;
  private dependencies: InputItem[] = [];
  private status: OutputStatus = OutputStatus.ORIG;
  private custom: string = "";

  constructor(data: OutputInfo) {
    super();
    // Coordinate information is represented in the graph.
    this.orig = data.orig;
    this.err = data.err;
  }

  /**
   * Adds a data dependency. Called only during initial graph construction.
   */
  public addDependency(dependency: InputItem): void {
    this.dependencies.push(dependency);
  }

  private valueChanged(): void {
    this.fireEvent(this.getValue());
  }

  public getType(): DDType { return DDType.OUTPUT; }
  public getValue(): string {
    switch (this.status) {
      case OutputStatus.ORIG:
        return this.orig;
      case OutputStatus.ERR:
        return this.err;
      case OutputStatus.CUSTOM:
        return this.custom;
      default:
        throw new Error("Invalid status: " + this.status);
    }
  }

  private changeStatus(status: OutputStatus): void {
    this.status = status;
    this.valueChanged();
  }

  public displayCustomValue(val: string): void {
    this.custom = val;
    this.changeStatus(OutputStatus.CUSTOM);
  }

  public displayError(): void {
    this.changeStatus(OutputStatus.ERR);
  }

  public displayOriginal(): void {
    this.changeStatus(OutputStatus.ORIG);
  }
}

/**
 * Object instantiation of an InputInfo from the JSON structure.
 */
class InputItem extends ChangeObservable implements DDItem {
  private orig: string;
  private err: string;
  private dependents: {
    noerr: string;
    output: OutputItem;
  }[] = [];
  private displayError: boolean = false;

  /**
   * Constructs an InputInfo object. Uses the data dependency graph to update
   * *outputs* with data dependencies.
   */
  constructor(graph: DataDependencyGraph, data: InputInfo) {
    super();
    var i: number, item: OutputItem;
    this.orig = data.orig;
    this.err = data.err;

    for (i = 0; i < data.outputs.length; i++) {
      // Grab each item, create two-way links.
      item = <OutputItem> graph.getItem(data.outputs[i]);
      assert(item.getType() === DDType.OUTPUT, "Input dependents must be outputs.");
      // Output -> input
      item.addDependency(this);
      // Input -> output
      this.dependents.push({
        noerr: data.outputs[i].noerr,
        output: item
      });
    }
  }

  public getType(): DDType { return DDType.INPUT; }
  public getValue(): string { return this.displayError ? this.err : this.orig; }
  private valueChanged(): void {
    this.fireEvent(this.getValue());
  }
}

/**
 * The data dependency graph. Produces a compact projection from three
 * dimensions (col, row, worksheet) to two dimensions (col, row).
 */
class DataDependencyGraph {
  /**
   * Three dimensional matrix. Cell items are at:
   * this.data[worksheet][x][y]
   */
  private data: { [worksheet: string]: DDItem[][] } = {};

  constructor(question: QuestionInfo) {
    // Add all items.

    // Outputs first; they have no dependents.
    var outputs = question.outputs,
      i: number, j: number;
    for (i = 0; i < outputs.length; i++) {
      this.addItem(outputs[i], new OutputItem(outputs[i]));
    }

    // Errors second.
    var errors = question.errors;
    for (i = 0; i < errors.length; i++) {
      // The InputItem constructor takes care of linking inputs/outputs
      // together.
      this.addItem(errors[i], new InputItem(this, errors[i]));
    }

    // @todo: The below things.
    // Compact each spreadsheet; eliminate whitespace.
    // 1. Collapse multiple empty rows into a single empty row.
    // 2. Collapse multiple empty columns into a single empty column.
    // Heuristics: Produce mapping from spreadsheets into a larger 2D array.
  }

  /**
   * Get the width of the 2D projection.
   */
  public getWidth(): number {
    var ws: string;
    for (ws in this.data) {
      // Finish.
    }
  }

  /**
   * Adds the given item to the graph.
   */
  public addItem(coord: SpreadsheetCoordinate, item: DDItem): void {
    var row = this.getRow(coord.worksheet, coord.x);

    if (row.length < coord.y) {
      this.data[coord.worksheet][coord.x] = row = row.concat(new Array(coord.y - row.length + 1));
    }
    row[coord.y] = item;
  }

  /**
   * Get the indicated worksheet.
   */
  public getWs(ws: string): DDItem[][] {
    var wsData: DDItem[][] = this.data[ws];
    if (typeof wsData === 'undefined') {
      throw new Error('Invalid worksheet: ' + ws);
    }
    return wsData;
  }

  /**
   * Get the specified row.
   */
  public getRow(ws: string, col: number): DDItem[] {
    var wsData = this.getWs(ws),
      colData = wsData[col];
    if (typeof colData === 'undefined') {
      throw new Error('Invalid row: ' + ws + ", " + col);
    }
    return colData;
  }

  /**
   * Retrieves the given coordinate from the spreadsheet, or throws an error
   * if not found.
   */
  public getItem(coord: SpreadsheetCoordinate): DDItem {
    var row = this.getRow(item.worksheet, coord.x), item = row[coord.y];
    assert(typeof item !== 'undefined', "Invalid coordinate: " + item.worksheet + ', ' + item.x + ', ' + item.y);
    return item;
  }
}

/**
 * Represents a single CheckCell task. Given a JSON object and a div id, it
 * will display the specified ranking question in the given div.
 */
class CheckCellQuestion {
  private graph: DataDependencyGraph;
  /**
   * @param data The JSON object with the question information.
   * @param divId The ID of the div where the question should be injected.
   */
  constructor(private data: QuestionInfo, divId: string) {
    this.graph = new DataDependencyGraph(data);

    var div: HTMLDivElement = <HTMLDivElement> document.getElementById(divId);
    div.appendChild(this.table);
  }
  
  /**
   * Constructs a data row of the table.
   */
  private constructRow(table: HTMLTableElement, row: string[], rowId: number) {
    var tr: HTMLTableRowElement = document.createElement('tr'),
      td: HTMLTableCellElement, i: number;

    td = document.createElement('td');
    td.innerText = "" + rowId;
    td.classList.add('header');
    tr.appendChild(td);
    for (i = 0; i < row.length; i++) {
      td = document.createElement('td');
      td.innerText = row[i];
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }

  /**
   * Constructs the <table> and its header.
   */
  private constructTable(): HTMLTableElement {
    var table: HTMLTableElement = document.createElement('table'),
      i: number, tr: HTMLTableRowElement = document.createElement('tr'),
      th: HTMLTableHeaderCellElement;

    // Construct header.
    th = document.createElement('th');
    th.classList.add('header');
    th.classList.add('rowHeaderHeader');
    tr.appendChild(th);
    for (i = 0; i < this.width; i++) {
      th = document.createElement('th');
      th.classList.add('header');
      th.innerText = getExcelColumn(i);
      tr.appendChild(th);
    }
    table.appendChild(tr);
    return table;
  }
}


window.onload = function () {
  var sampleTable = new CheckCellQuestion(sampleQuestion, 'sample');
};
