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
          "worksheet": "sheet2",
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
          "worksheet": "sheet2",
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
      "worksheet": "sheet2",
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
  addEventListener(type: string, cb: (data: DDItem) => void): void;
  /**
   * Get the current value of the item.
   */
  getValue(): string;
  /**
   * Is the current value erroneous? Determines if the item should be
   * highlighted or not.
   */
  isValueErroneous(): boolean;
}

class ChangeObservable<T> {
  constructor(events: string[]) {
    var i: number;
    for (i = 0; i < events.length; i++) {
      this.listeners[events[i]] = [];
    }
  }

  private listeners: { [event: string]: { (data: T): void }[] } = {};

  public addEventListener(type: string, cb: (data: T) => void): void {
    if (this.listeners.hasOwnProperty(type)) {
      this.listeners[type].push(cb);
    }
  }

  public fireEvent(event: string, data: T) {
    var i: number;
    for (i = 0; i < this.listeners[event].length; i++) {
      this.listeners[event][i](data);
    }
  }
}

enum OutputStatus { ORIG, ERR, CUSTOM }

/**
 * Object instantiation of an OutputInfo from the JSON structure.
 */
class OutputItem extends ChangeObservable<OutputItem> implements DDItem {
  private orig: string;
  private err: string;
  private dependencies: InputItem[] = [];
  private status: OutputStatus = OutputStatus.ORIG;
  private custom: string = "";

  constructor(data: OutputInfo) {
    super(['changed']);
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
    this.fireEvent('changed', this);
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

  private changeStatus(status: OutputStatus, val?: string): void {
    var valChanged: boolean = false;
    if (this.status !== status) {
      valChanged = true;
      this.status = status;
    }
    if (status === OutputStatus.CUSTOM && this.custom !== val) {
      valChanged = true;
      this.custom = val;
    }
    if (valChanged) {
      this.valueChanged();
    }
  }

  public displayCustomValue(val: string): void {
    this.changeStatus(OutputStatus.CUSTOM, val);
  }

  public displayError(): void {
    this.changeStatus(OutputStatus.ERR);
  }

  public displayOriginal(): void {
    this.changeStatus(OutputStatus.ORIG);
  }

  public isValueErroneous(): boolean {
    return this.status !== OutputStatus.ORIG;
  }
}

/**
 * Object instantiation of an InputInfo from the JSON structure.
 */
class InputItem extends ChangeObservable<InputItem> implements DDItem {
  private orig: string;
  private err: string;
  private dependents: {
    noerr: string;
    output: OutputItem;
  }[] = [];
  private shouldDisplayError: boolean = false;

  /**
   * Constructs an InputInfo object. Uses the data dependency graph to update
   * *outputs* with data dependencies.
   */
  constructor(graph: DataDependencyGraph, data: InputInfo) {
    super(['changed']);
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
  public getValue(): string { return this.shouldDisplayError ? this.err : this.orig; }
  private valueChanged(): void {
    this.fireEvent('changed', this);
  }

  public isValueErroneous(): boolean {
    return this.shouldDisplayError;
  }

  /**
   * Changes the value of this input item, and, optionally, its dependents.
   * @param toError Are we displaying an error, or the original value?
   * @param singleError Are we displaying this input item's error alone (true),
   *   or are we showing all input errors at once (false)?
   */
  public changeValue(toError: boolean, singleError?: boolean): void {
    var i: number, savedShouldDisplayError = this.shouldDisplayError;
    if (toError) {
      this.shouldDisplayError = true;
      if (singleError) {
        for (i = 0; i < this.dependents.length; i++) {
          // @todo Need to ask Dan about this value. I don't understand why it's
          // 'noerr' instead of 'err'.
          this.dependents[i].output.displayCustomValue(this.dependents[i].noerr);
        }
      }
    } else {
      // No error.
      this.shouldDisplayError = false;
      for (i = 0; i < this.dependents.length; i++) {
        this.dependents[i].output.displayOriginal();
      }
    }

    if (this.shouldDisplayError !== savedShouldDisplayError) {
      this.valueChanged();
    }
  }
}

/**
 * The data dependency graph.
 */
class DataDependencyGraph {
  /**
   * Three dimensional matrix. Cell items are at:
   * this.data[worksheet][x][y]
   */
  private data: { [worksheet: string]: DDItem[][] } = {};
  private inputs: InputItem[] = [];
  private outputs: OutputItem[] = [];

  constructor(question: QuestionInfo) {
    // Add all items.

    // Outputs first; they have no dependents.
    var outputs = question.outputs,
      i: number, j: number;
    for (i = 0; i < outputs.length; i++) {
      var outputItem = new OutputItem(outputs[i]);
      this.addItem(outputs[i], outputItem);
      this.outputs.push(outputItem);
    }

    // Errors second.
    var errors = question.errors;
    for (i = 0; i < errors.length; i++) {
      // The InputItem constructor takes care of linking inputs/outputs
      // together.
      var item: InputItem = new InputItem(this, errors[i]);
      this.addItem(errors[i], item);
      this.inputs.push(item);
    }

    // @todo: The below things.
    // Compact each spreadsheet; eliminate whitespace.
    // 1. Collapse multiple empty rows into a single empty row.
    // 2. Collapse multiple empty columns into a single empty column.
  }

  /**
   * Get the maximum row width.
   */
  public getWidth(): number {
    var ws: string, wsData: DDItem[][], row: DDItem[], i: number,
      width: number = 0;
    for (ws in this.data) {
      if (this.data.hasOwnProperty(ws)) {
        wsData = this.data[ws];
        for (i = 0; i < wsData.length; i++) {
          row = wsData[i];
          if (row.length > width) {
            width = row.length;
          }
        }
      }
    }
    return width;
  }

  /**
   * Get the maximum column height.
   */
  public getHeight(): number {
    var ws: string, wsData: DDItem[][], height: number = 0;
    for (ws in this.data) {
      if (this.data.hasOwnProperty(ws)) {
        wsData = this.data[ws];
        if (wsData.length > height) {
          height = wsData.length;
        }
      }
    }
    return height;
  }

  /**
   * ABSTRACTIONS!
   */
  public getData(): { [worksheet: string]: DDItem[][] } {
    return this.data;
  }

  /**
   * Adds the given item to the graph.
   */
  public addItem(coord: SpreadsheetCoordinate, item: DDItem): void {
    // Does the WS exist?
    var wsData = this.data[coord.worksheet];
    if (!wsData) {
      wsData = this.data[coord.worksheet] = [];
    }

    // Does the row exist?
    var row = wsData[coord.x];
    if (!row) {
      var width = wsData.length, i: number;
      for (i = width; i <= coord.x; i++) {
        wsData[i] = [];
      }
      row = wsData[coord.x] = [];
    }

    if (row.length < coord.y) {
      // Pad with empty entries.
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
    var row = this.getRow(coord.worksheet, coord.x), item = row[coord.y];
    assert(typeof item !== 'undefined', "Invalid coordinate: " + coord.worksheet + ', ' + coord.x + ', ' + coord.y);
    return item;
  }

  public getInputs(): InputItem[] {
    return this.inputs;
  }

  public getOutputs(): OutputItem[] {
    return this.outputs;
  }
}

/**
 * Represents a single worksheet within a CheckCellQuestion. Each question may
 * have multiple worksheets.
 */
class WorksheetTable extends ChangeObservable<WorksheetTable> {
  private containsErrors: boolean = false;
  private tableDiv: HTMLDivElement = document.createElement('div');
  private errorCount: number = 0;

  /**
   * @param data The data to display in the worksheet.
   * @param width The width of the table. Specified here so all worksheets are
   *   equal widths.
   * @param height The height of the table. Specified here so all worksheets
   *   are equal heights.
   * @param question The CheckCellQuestion that this worksheet belongs to.
   */
  constructor(private question: CheckCellQuestion, private name: string, private data: DDItem[][],
    private width: number, private height: number) {
    super(['changed']);
      this.tableDiv = document.createElement('div');
      this.tableDiv.appendChild(this.constructTable());
      this.tableDiv.classList.add('tabbertab');
      this.tableDiv.setAttribute('title', this.name);
  }

  public getName(): string { return this.name; }
  public isDisplayingErrors(): boolean {
    return this.errorCount > 0;
  }

  public getDiv(): HTMLDivElement {
    return this.tableDiv;
  }

  private errorDelta(change: number) {
    // We use errorCount as a semaphore. When it goes to 0, the worksheet no
    // longer contains errors. When it goes above 0, the worksheet now has
    // errors.

    // Add *prior* to triggering events so callbacks get correct
    // 'isDisplayingErrors' value.
    this.errorCount += change;
    if ((this.errorCount - change) === 0) {
      this.fireEvent('changed', this);
    } else if (this.errorCount === 0) {
      this.fireEvent('changed', this);
    }

    // Invariant: Error count never goes negative.
    assert(this.errorCount >= 0);
  }

  private constructCell(data: DDItem): HTMLTableCellElement {
    var cell: HTMLTableCellElement = document.createElement('td');
    cell.addEventListener('click', (ev) => {
      // Check if this is an erroneous value, and store that info.
      var currentlyError: boolean = data.isValueErroneous();
      // Clear *all* errors.
      this.question.clearErrors();
      // Change value appropriately using the stored information.
      if (!currentlyError && data.getType() === DDType.INPUT) {
        // The clicked cell is an input cell, and is was not in an error state.
        // Change it into an error, and update its dependents.
        var input: InputItem = <InputItem> data;
        input.changeValue(true, true);
      }
    });

    // Change events can be triggered by the global spreadsheet, *or* by the
    // above click handler.
    data.addEventListener('changed', (data: DDItem) => {
      // Update displayed value.
      cell.innerText = data.getValue();
      // Fire event if this is an error.
      if (data.isValueErroneous()) {
        // Add the 'erroneous' style.
        cell.classList.add('erroneous');
        this.errorDelta(1);
      } else {
        // Remove the 'erroneous' style.
        cell.classList.remove('erroneous');
        this.errorDelta(-1);
      }
    });
    cell.innerText = data.getValue();
    return cell;
  }

  /**
   * Constructs a data row of the table.
   */
  private constructRow(row: DDItem[], rowId: number): HTMLTableRowElement {
    var tr: HTMLTableRowElement = document.createElement('tr'),
      td: HTMLTableCellElement, i: number, item: DDItem;

    td = document.createElement('td');
    td.innerText = "" + rowId;
    td.classList.add('header');
    tr.appendChild(td);
    for (i = 0; i < this.width; i++) {
      item = row[i];
      if (typeof item === 'undefined') {
        tr.appendChild(document.createElement('td'));
      } else {
        tr.appendChild(this.constructCell(item));
      }
    }
    return tr;
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
    for (i = 0; i < this.height; i++) {
      if (i < this.data.length) {
        table.appendChild(this.constructRow(this.data[i], i+1));
      } else {
        table.appendChild(this.constructRow([], i+1));
      }
    }
    return table;
  }
}

/**
 * Represents a single CheckCell question. Given a JSON object and a div id, it
 * will display the specified ranking question in the given div.
 * @todo track active cells through spreadsheets, deactivate when needed
 * @todo track/handle switching active worksheets.
 */
class CheckCellQuestion {
  private graph: DataDependencyGraph;
  private parentDiv: HTMLDivElement;
  private divElement: HTMLDivElement;
  private globalErrors: boolean = false;
  private tables: { [ws: string]: WorksheetTable } = {};

  /**
   * @param data The JSON object with the question information.
   * @param divId The ID of the div where the question should be injected.
   */
  constructor(private data: QuestionInfo, divId: string) {
    this.graph = new DataDependencyGraph(data);
    this.divElement = document.createElement('div');
    this.divElement.classList.add('tabber');

    var graphData = this.graph.getData(), i: number, ws: string,
      width: number = this.graph.getWidth(), height: number = this.graph.getHeight();
    for (ws in graphData) {
      if (graphData.hasOwnProperty(ws)) {
        var wsTable = new WorksheetTable(this, ws, graphData[ws], width, height);
        this.tables[ws] = wsTable;
        wsTable.addEventListener('changed', (data: WorksheetTable): void => {
          this.toggleTabError(data.getName(), data.isDisplayingErrors());
        });
        this.divElement.appendChild(wsTable.getDiv());
      }
    }

    this.parentDiv = <HTMLDivElement> document.getElementById(divId);
    this.parentDiv.appendChild(this.divElement);

    // Button to toggle all errors.
    var vladimirButin = document.createElement('button');
    vladimirButin.innerText = "Toggle all errors";
    vladimirButin.addEventListener('click', (ev): void => {
      this.toggleAllErrors(!this.globalErrors);
    });
    this.parentDiv.appendChild(vladimirButin);
  }

  public toggleTabError(tabName: string, isError: boolean): void {
    // Get the tab element for the worksheet.
    var tabDiv = this.parentDiv.getElementsByClassName('tabberlive');
    assert(tabDiv.length === 1);
    var tabList: HTMLUListElement = <HTMLUListElement> tabDiv[0].childNodes[0];

    var children = tabList.children, i: number;
    for (i = 0; i < children.length; i++) {
      var child: HTMLUListElement = <HTMLUListElement> children[i],
        tabLink: HTMLAnchorElement = <HTMLAnchorElement> child.children[0];
      if (child.innerText === tabName) {
        if (isError) {
          tabLink.classList.add('errorTab');
        } else {
          tabLink.classList.remove('errorTab');
        }
        return;
      }
    }
    assert(false, "Couldn't find tab " + tabName);
  }

  public toggleAllErrors(enable: boolean): void {
    if (enable !== this.globalErrors) {
      this.clearErrors();
      this.globalErrors = enable;
      var inputs = this.graph.getInputs(),
        outputs = this.graph.getOutputs(),
        i: number;
      for (i = 0; i < inputs.length; i++) {
        inputs[i].changeValue(this.globalErrors, false);
      }
      for (i = 0; i < outputs.length; i++) {
        enable ? outputs[i].displayError() : outputs[i].displayOriginal();
      }
    }
  }

  public clearErrors(): void {
    var inputs = this.graph.getInputs(), i: number;
    for (i = 0; i < inputs.length; i++) {
      if (inputs[i].isValueErroneous()) {
        inputs[i].changeValue(false);
      }
    }
  }

  public allErrorsEnabled(): boolean {
    return this.globalErrors;
  }
}


declare var tabberAutomatic: Function;
window.onload = function () {
  var sampleTable = new CheckCellQuestion(sampleQuestion, 'sample');
  tabberAutomatic();
};

// Tabber options
window['tabberOptions'] = { manualStartup: true };
