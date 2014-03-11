/// <reference path="ref/jquery.d.ts" />
/// <reference path="ref/jqueryui.d.ts" />
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

interface CellInfo extends SpreadsheetCoordinate {
  orig: string;
  err: string;
}

/**
 * Represents spreadsheet *outputs*.
 */
interface OutputInfo extends CellInfo {
  formula: string;
}

/**
 * Represents spreadsheet *inputs*.
 */
interface InputInfo extends CellInfo {
  outputs: {
    x: number;
    y: number;
    worksheet: string;
    noerr: string;
  }[];
  style: CellStyle;
}

interface CellStyle {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  'font-face': string;
  'font-size': number;
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
      ],
      "style": {
        "bold": true,
        "italic": false,
        "underline": false,
        "font-face": "Arial",
        "font-size": 12
      }
    },
    {
      "x": 1,
      "y": 2,
      "worksheet": "fun",
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
      ],
      "style": {
        "bold": false,
        "italic": false,
        "underline": false,
        "font-face": "Wingdings",
        "font-size": 11
      }
    },
    {
      "x": 1,
      "y": 3,
      "worksheet": "fun",
      "orig": "10.0",
      "err": "10.0",
      "outputs": [],
      "style": {
        "bold": false,
        "italic": false,
        "underline": false,
        "font-face": "Wingdings",
        "font-size": 11
      }
    }
  ],
  "outputs": [
    {
      "x": 6,
      "y": 7,
      "worksheet": "sheet1",
      "orig": "0.0",
      "err": "100",
      "formula": "=SUM(A1:A10)"
    },
    {
      "x": 7,
      "y": 7,
      "worksheet": "sheet1",
      "orig": "3.14159265359",
      "err": "99.0",
      "formula": "=AVERAGE(Z22:Z23)"
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

var __id = 0;
function nextId(): number {
  return __id++;
}

function coords2string(coords: SpreadsheetCoordinate): string {
  return coords.worksheet + " " + getExcelColumn(coords.y - 1) + (coords.x);
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
   * Fires the specified event on the object.
   */
  fireEvent(type: string, data: DDItem): void;
  /**
   * Get the current value of the item.
   */
  getValue(): string;
  /**
   * Is the current value erroneous? Determines if the item should be
   * highlighted or not.
   */
  isValueErroneous(): boolean;
  getCoords(): SpreadsheetCoordinate;
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
  private dependencies: InputItem[] = [];
  private status: OutputStatus = OutputStatus.ERR;
  private custom: string = "";

  constructor(private data: OutputInfo) {
    super(['changed']);
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
        return this.data.orig;
      case OutputStatus.ERR:
        return this.data.err;
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
    return this.status === OutputStatus.ERR;
  }

  public getCoords(): SpreadsheetCoordinate {
    return {
      worksheet: this.data.worksheet,
      x: this.data.x,
      y: this.data.y
    };
  }

  public getFormula(): string {
    return this.data.formula;
  }
}

/**
 * Object instantiation of an InputInfo from the JSON structure.
 */
class InputItem extends ChangeObservable<InputItem> implements DDItem {
  private dependents: {
    noerr: string;
    output: OutputItem;
  }[] = [];
  private shouldDisplayError: boolean = true;
  private draggable: boolean = true;

  /**
   * Constructs an InputInfo object. Uses the data dependency graph to update
   * *outputs* with data dependencies.
   */
  constructor(graph: DataDependencyGraph, private data: InputInfo) {
    super(['changed', 'draggableChanged']);
    var i: number, item: OutputItem;

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

    if (this.isContext()) {
      this.shouldDisplayError = false;
    }
  }

  public setDraggable(draggable: boolean) {
    if (draggable !== this.draggable) {
      this.draggable = draggable;
      this.fireEvent('draggableChanged', this);
    }
  }

  public isDraggable(): boolean {
    return this.draggable;
  }

  public getType(): DDType { return DDType.INPUT; }
  public getValue(): string { return this.shouldDisplayError ? this.data.err : this.data.orig; }
  private valueChanged(): void {
    this.fireEvent('changed', this);
  }

  public isValueErroneous(): boolean {
    return this.shouldDisplayError;
  }

  /**
   * Changes the value of this input item.
   */
  public changeValue(toError: boolean): void {
    if (this.shouldDisplayError !== toError) {
      this.shouldDisplayError = toError;
      this.valueChanged();
    }
  }

  public getDependents(): { noerr: string; output: OutputItem; }[] {
    return this.dependents.slice(0);
  }

  public getCoords(): SpreadsheetCoordinate {
    return {
      worksheet: this.data.worksheet,
      x: this.data.x,
      y: this.data.y
    };
  }

  /**
   * Is this item a context item, e.g. it's not an item that needs to be ranked?
   */
  public isContext(): boolean {
    return this.data.orig === this.data.err || this.dependents.length === 0;
  }
  
  public getStyle(): CellStyle {
    return this.data.style;
  }

  public getErrorValue(): string {
    return this.data.err;
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
      if (!item.isContext()) {
        this.inputs.push(item);
      }
    }
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
class WorksheetTable {
  private containsErrors: boolean = false;
  private tableDiv: JQuery;
  private errorCount: number = 0;
  private tabHighlighted: boolean = false;

  /**
   * @param data The data to display in the worksheet.
   * @param width The width of the table. Specified here so all worksheets are
   *   equal widths.
   * @param height The height of the table. Specified here so all worksheets
   *   are equal heights.
   * @param question The CheckCellQuestion that this worksheet belongs to.
   */
  constructor(private question: CheckCellQuestion, private name: string, private tabAnchor: JQuery, private data: DDItem[][],
    private width: number, private height: number) {
      this.tableDiv = $('<div>')
        .addClass('ccTab')
        .attr('id', tabAnchor.attr('href').slice(1))
        .attr('title', this.name)
        .append(this.constructTable());
  }

  public getName(): string { return this.name; }
  public isDisplayingErrors(): boolean {
    return this.errorCount > 0;
  }

  public getDiv(): JQuery {
    return this.tableDiv;
  }

  public toggleHighlighting(enable: boolean) {
    if (enable !== this.tabHighlighted) {
      this.tabHighlighted = enable;
      if (enable) {
        this.tabAnchor.text(this.name + '*').addClass('ccWsTabChange');
      } else {
        this.tabAnchor.text(this.name).removeClass('ccWsTabChange');
      }
    }
  }

  private constructBlankCell(): JQuery {
    return $('<td>');
  }

  private constructCell(data: DDItem): JQuery {
    var cell: JQuery = this.constructBlankCell();
    if (data.getType() === DDType.INPUT) {
      cell.addClass('ccInput');
    } else {
      cell.addClass('ccOutput');
    }

    // Only listen for clicks and drags on non-context input cells.
    if (data.getType() === DDType.INPUT) {
      var input: InputItem = <InputItem> data;
      if (input.isContext()) {
        // Properly item style.
        var style = input.getStyle();
        if (style.bold) {
          cell.css('font-weight', 'bold');
        }
        if (style.underline) {
          cell.css('text-decoration', 'underline');
        }
        if (style.italic) {
          cell.css('font-style', 'italic');
        }
        cell.css('font-family', style['font-face']);
        cell.css('font-size', '' + style['font-size'] + 'pt')
      } else {
        cell.on('click', (ev) => {
          if (data.isValueErroneous()) {
            // Input item is erroneous and the user clicked on it.
            // Transition to a state where it is not erroneous.
            this.question.changeStatus(SpreadsheetStatus.ALL_BUT_ONE_ERROR, <InputItem> data);
          } else if (this.question.getStatus() === SpreadsheetStatus.ALL_BUT_ONE_ERROR) {
            // Input item is not erroneous, it is the only item not erroneous,
            // and the user clicked on it. Transition to a state where it is
            // erroneous.
            this.question.changeStatus(SpreadsheetStatus.ALL_ERRORS);
          }
          // Ignore clicks when all errors are off.
        });

        data.addEventListener('draggableChanged', (data: DDItem) => {
          if (input.isDraggable()) {
            cell.removeClass('ccInputUndraggable')
              .addClass('ccInputDraggable')
              .draggable({
                cursor: 'move',
                revert: 'invalid',
                helper: () => {
                  return $('<li>' + coords2string(data.getCoords()) + ': ' + input.getErrorValue() + '</li>').addClass('ccListItem').data("DDItem", data);
                }
              });
          } else {
            cell.removeClass('ccInputDraggable')
              .addClass('ccInputUndraggable')
              .draggable('destroy');
          }
        });

        // Bootstrap.
        data.fireEvent('draggableChanged', data);
      }
    } else {
      var output: OutputItem = <OutputItem> data;
      // title === alt text for non-images.
      cell.attr('title', output.getFormula());
    }

    // Change events can be triggered by the global spreadsheet, *or* by the
    // above click handler.
    data.addEventListener('changed', (data: DDItem) => {
      // Update displayed value.
      cell.text(data.getValue());
      if (data.getType() === DDType.INPUT) {
        if (data.isValueErroneous()) {
          // Add the 'erroneous' style.
          cell.addClass('ccInputError');
        } else {
          // Remove the 'erroneous' style.
          cell.removeClass('ccInputError');
        }
      } else {
        // Check if we are in NO_ERROR or ALL_BUT_ONE_ERROR.
        if (this.question.getStatus() !== SpreadsheetStatus.ALL_ERRORS) {
          cell.addClass('ccOutputChange');
        } else {
          cell.removeClass('ccOutputChange');
        }
      }

      // Highlight our tab if this element is part of a single disabled
      // error.
      this.toggleHighlighting((!data.isValueErroneous()) && this.question.getStatus() === SpreadsheetStatus.ALL_BUT_ONE_ERROR);
    });
    // Bootstrap cell value.
    data.fireEvent('changed', data);
    return cell;
  }

  /**
   * Constructs a data row of the table.
   */
  private constructRow(row: DDItem[], rowId: number): JQuery {
    var tr: JQuery = $('<tr>'), i: number, item: DDItem;

    tr.append(this.constructBlankCell().text("" + (rowId - 1)).addClass('ccHeader'));
    // XXX: Excel is 1-indexed. Ignore the 0th cell.
    for (i = 1; i < this.width; i++) {
      item = row[i];
      if (typeof item === 'undefined') {
        tr.append(this.constructBlankCell());
      } else {
        tr.append(this.constructCell(item));
      }
    }
    return tr;
  }

  /**
   * Constructs the <table> and its header.
   */
  private constructTable(): JQuery {
    var table: JQuery,
      i: number, tr: JQuery = $('<tr>');

    table = $('<table>').addClass('ccMain');
    // Construct header.
    tr.append($('<th>')
      .addClass('ccHeader')
      .addClass('ccRowHeaderHeader')
    );
    // XXX: Excel is 1-indexed.
    for (i = 1; i < this.width; i++) {
      tr.append($('<th>')
        .addClass('ccHeader')
        .text(getExcelColumn(i - 1))
      );
    }
    table.append(tr);
    // XXX: Excel is 1-indexed.
    for (i = 1; i < this.height; i++) {
      if (i < this.data.length) {
        table.append(this.constructRow(this.data[i], i+1));
      } else {
        table.append(this.constructRow([], i+1));
      }
    }
    return table;
  }
}

/**
 * The spreadsheet can be in one of three states:
 * - All errors on (ALL_ERRORS)
 * - No errors on (NO_ERRORS)
 * - All but one error on (ALL_BUT_ONE_ERROR)
 */
enum SpreadsheetStatus {
  ALL_ERRORS, NO_ERRORS, ALL_BUT_ONE_ERROR
}

/**
 * Represents a single CheckCell question. Given a JSON object and a div id, it
 * will display the specified ranking question in the given div.
 * @todo track active cells through spreadsheets, deactivate when needed
 * @todo track/handle switching active worksheets.
 */
class CheckCellQuestion {
  private graph: DataDependencyGraph;
  private parentDiv: JQuery;
  private questionDiv: JQuery;
  private status: SpreadsheetStatus = SpreadsheetStatus.ALL_ERRORS;
  private tables: { [ws: string]: WorksheetTable } = {};
  private disabledError: InputItem = null;
  private toggleButton: JQuery;

  // Used for drag n' drop events.
  private rankListDiv: JQuery;
  private unimportantListDiv: JQuery;

  /**
   * @param data The JSON object with the question information.
   * @param divId The ID of the div where the question should be injected.
   */
  constructor(private data: QuestionInfo, private divId: string) {
    this.graph = new DataDependencyGraph(data);
    this.questionDiv = $('<div>').addClass('ccQuestionDiv').attr('id', 'ccQuestion' + nextId());

    var graphData = this.graph.getData(), i: number, ws: string,
      width: number = this.graph.getWidth(),
      height: number = this.graph.getHeight(), tab: JQuery,
      wsTable: WorksheetTable, tabList: JQuery = $('<ul>').addClass('ccTabList');

    this.questionDiv.append(tabList);
    for (ws in graphData) {
      if (graphData.hasOwnProperty(ws)) {
        // Create tab.
        tab = $('<li>')
          .append($('<a>')
            .attr('href', '#ccTab' + nextId())
            .text(ws)
          );
        tabList.append(tab);
        // Create body.
        // NOTE: WST wants the anchor, not the list element.
        wsTable = new WorksheetTable(this, ws, $(tab.find('a')[0]), graphData[ws], width, height);
        this.tables[ws] = wsTable;
        this.questionDiv.append(wsTable.getDiv());
      }
    }

    this.parentDiv = $('#' + divId).append(this.questionDiv);
    // Enable tabs.
    this.questionDiv.tabs();

    // Ranking table.
    var sharedListClass = 'dragList' + nextId(), self = this,
      ul = $('<ul>')
        .addClass(sharedListClass)
        .sortable({
          revert: 'false',
          connectWith: '.' + sharedListClass
        }),
      dropHandler = function (e, ui) {
        // Only append if this is a child element of the question div.
        if ($(ui.draggable).closest('#' + self.questionDiv.attr('id')).length > 0) {
          var item: InputItem = ui.helper.data('DDItem'),
            helper: JQuery = ui.helper;
          $($(this).find('ul')[0]).append($('<li>').text(helper.text()).addClass('ccListItem').data('DDItem', item));
          // Wait one turn for jQuery UI to do it's thing before we disable
          // dragging.
          setTimeout(() => { item.setDraggable(false); }, 0);
        }
      };
    this.rankListDiv = $('<div>')
      .addClass('ccRankList')
      .append($('<h4>Ranked List</h4>'))
      .append(ul)
      .droppable({
        tolerance: 'pointer',
        accept: () => { return true; },
        drop: dropHandler
      });
    this.parentDiv.append(this.rankListDiv);

    // Unimportant table.
    this.unimportantListDiv = $('<div>')
      .addClass('ccUnimportantList')
      .droppable({
        tolerance: 'pointer',
        accept: () => { return true; },
        drop: dropHandler
      })
      .append($('<h4>Unimportant Inputs</h4>'))
      .append($('<ul>')
        .addClass(sharedListClass)
        .sortable({
          revert: 'false',
          connectWith: '.' + sharedListClass
        })
      );
    this.parentDiv.append(this.unimportantListDiv);

    // Button to toggle all errors.
    this.toggleButton = $('<button>')
      .text("Toggle errors off")
      .on('click', (ev): void => {
        if (this.toggleButton.text() === 'Toggle errors off') {
          this.changeStatus(SpreadsheetStatus.NO_ERRORS);
          this.toggleButton.text("Toggle errors on");
        } else {
          this.changeStatus(SpreadsheetStatus.ALL_ERRORS);
          this.toggleButton.text("Toggle errors off");
        }
      });
    this.parentDiv.append($('<br>')).append(this.toggleButton);

    // Validate button.
    var validateBtn = $('<button>')
      .text('Validate')
      .on('click', (ev): void => {
        try {
          var ranking = this.getRanking();
          alert("Validates: " + JSON.stringify(ranking));
        } catch (e) {
          alert("Does not validate: " + e.toString());
        }
      });
    this.parentDiv.append(validateBtn);
  }

  public getRanking(): { unimportant: SpreadsheetCoordinate[]; ranking: SpreadsheetCoordinate[] } {
    var inputs: InputItem[] = this.graph.getInputs(), i: number, item: InputItem,
      rv: {
        unimportant: SpreadsheetCoordinate[];
        ranking: SpreadsheetCoordinate[];
      } = { unimportant: [], ranking: [] },
      coords2item: { [coords: string]: InputItem } = {}, coords: string;

    // Hash from coords => item
    for (i = 0; i < inputs.length; i++) {
      coords2item[coords2string(inputs[i].getCoords())] = inputs[i];
    }

    // Find each item in the unimportant list in the hash.
    var unimportantList = this.unimportantListDiv.find('ul').children();
    for (i = 0; i < unimportantList.length; i++) {
      // XXX: Hack cuz list values are "coords: value".
      coords = $(unimportantList[i]).text().split(':')[0];
      item = coords2item[coords];
      delete coords2item[coords];
      assert(typeof item !== 'undefined');
      rv.unimportant.push(item.getCoords());
    }

    // Find each item in the rank list in the hash.
    var rankList = this.rankListDiv.find('ul').children();
    for (i = 0; i < rankList.length; i++) {
      // XXX: Hack cuz list values are "coords: value".
      coords = $(rankList[i]).text().split(':')[0];
      item = coords2item[coords];
      delete coords2item[coords];
      assert(typeof item !== 'undefined');
      rv.ranking.push(item.getCoords());
    }

    // Throw an error if the hash is not empty.
    if (Object.keys(coords2item).length > 0) {
      throw new Error('The following items have not been ranked: ' + JSON.stringify(Object.keys(coords2item)));
    }

    return rv;
  }

  public getDivId(): string { return this.divId; }

  private getWorksheetTab(wsName: string): JQuery {
    var tabDiv = this.parentDiv.find('.tabberlive'),
      tab = tabDiv.find("a:contains('" + wsName + "')");
    if (tab.length > 0) {
      assert(tab.length === 1);
      return $(tab[0]);
    } else {
      tab = tabDiv.find("a:contains('" + wsName + "*')");
      assert(tab.length === 1);
      return $(tab[0]);
    }
  }

  /**
   * Change the global status of the spreadsheet.
   */
  public changeStatus(status: SpreadsheetStatus, item?: InputItem): void {
    // Sanity checks.
    if (status === SpreadsheetStatus.ALL_BUT_ONE_ERROR && item == null) {
      throw new Error("Item must be specified.");
    } else if (status !== SpreadsheetStatus.ALL_BUT_ONE_ERROR &&
      status !== SpreadsheetStatus.ALL_ERRORS && status !== SpreadsheetStatus.NO_ERRORS) {
        throw new Error("Invalid status: " + status);
    }

    // Transition table time!
    var i: number, oldStatus: SpreadsheetStatus = this.status,
      inputs: InputItem[], outputs: OutputItem[], toError: boolean;
    this.status = status;
    this.disabledError = null;
    switch (status) {
      case SpreadsheetStatus.NO_ERRORS:
        // INTENTIONAL FALL-THRU
      case SpreadsheetStatus.ALL_ERRORS:
        // Just update everything to the correct state.
        inputs = this.graph.getInputs();
        outputs = this.graph.getOutputs();
        toError = status === SpreadsheetStatus.ALL_ERRORS;
        for (i = 0; i < inputs.length; i++) {
          inputs[i].changeValue(toError);
        }
        for (i = 0; i < outputs.length; i++) {
          toError ? outputs[i].displayError() : outputs[i].displayOriginal();
        }
        break;
      case SpreadsheetStatus.ALL_BUT_ONE_ERROR:
        var dependents = item.getDependents();
        // Change to the ALL_ERROR case, if not already.
        if (oldStatus !== SpreadsheetStatus.ALL_ERRORS) {
          this.changeStatus(SpreadsheetStatus.ALL_ERRORS);
          this.status = status;
        }
        item.changeValue(false);
        for (i = 0; i < dependents.length; i++) {
          dependents[i].output.displayCustomValue(dependents[i].noerr);
        }
        break;
    }
    return;
  }

  public getStatus(): SpreadsheetStatus {
    return this.status;
  }
}

window.onload = function () {
  var sampleTable = new CheckCellQuestion(sampleQuestion, 'sample');
  // new CheckCellQuestion(sampleQuestion, 'sample2');
};

