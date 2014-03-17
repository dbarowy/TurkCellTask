/// <reference path="ref/jquery.d.ts" />
/// <reference path="ref/jqueryui.d.ts" />
/**
 * Client-side logic for laying out a web page for a CheckCell error-ranking
 * question. Consumes a JSON blob and spits out a question in the desired div.
 */

// #region Dumb Browser Polyfills
// From https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object/keys
if (!Object.keys) {
  Object.keys = (function () {
    var hasOwnProperty = Object.prototype.hasOwnProperty,
      hasDontEnumBug = !({ toString: null }).propertyIsEnumerable('toString'),
      dontEnums = [
        'toString',
        'toLocaleString',
        'valueOf',
        'hasOwnProperty',
        'isPrototypeOf',
        'propertyIsEnumerable',
        'constructor'
      ],
      dontEnumsLength = dontEnums.length;

    return function (obj) {
      if (typeof obj !== 'object' && (typeof obj !== 'function' || obj === null)) {
        throw new TypeError('Object.keys called on non-object');
      }

      var result = [], prop, i;

      for (prop in obj) {
        if (hasOwnProperty.call(obj, prop)) {
          result.push(prop);
        }
      }

      if (hasDontEnumBug) {
        for (i = 0; i < dontEnumsLength; i++) {
          if (hasOwnProperty.call(obj, dontEnums[i])) {
            result.push(dontEnums[i]);
          }
        }
      }
      return result;
    };
  } ());
}

if (!Array.prototype.indexOf) {
  Array.prototype.indexOf = function (searchElement, fromIndex?) {
    if (this === undefined || this === null) {
      throw new TypeError('"this" is null or not defined');
    }

    var length = this.length >>> 0; // Hack to convert object.length to a UInt32

    fromIndex = +fromIndex || 0;

    if (Math.abs(fromIndex) === Infinity) {
      fromIndex = 0;
    }

    if (fromIndex < 0) {
      fromIndex += length;
      if (fromIndex < 0) {
        fromIndex = 0;
      }
    }

    for (; fromIndex < length; fromIndex++) {
      if (this[fromIndex] === searchElement) {
        return fromIndex;
      }
    }

    return -1;
  };
}

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
  output: {
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

// #endregion JSON Typings

// #region Helper Functions

/**
 * Converts a number into an Excel column ID.
 * e.g. 1 => A, 27 => AA.
 */
function getExcelColumn(i: number): string {
  var chars: string[] = [];
  // Excel is 1-based, algorithm is 0-based.
  i--;
  do {
    chars.push(String.fromCharCode(65 + (i % 26)));
    i = Math.floor(i / 26) - 1;
  } while (i > -1);
  chars.reverse();
  return chars.join('');
}

function getColClass(wsClassID: string, col: number): string {
  return wsClassID + 'Col' + getExcelColumn(col); 
}

function getRowClass(wsClassID: string, row: number): string {
  return wsClassID + 'Row' + row; 
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
  return coords.worksheet + " " + getExcelColumn(coords.x) + (coords.y);
}

/**
 * Given two table cells, make dest the same outer width as src.
 */
function matchCellWidth(src: JQuery, dest: JQuery) {
  var srcOuterWidth: number = src.outerWidth(),
    destOuterWidth: number;

  dest.css({
    'min-width': srcOuterWidth
  });

  // After applying that change, how wide is the destination cell?
  destOuterWidth = dest.outerWidth();

  // Box model, how does it work?!
  // The truth is, it doesn't matter. Try something, see if we come up narrow
  // (or wide), and then adjust to the current browser's behavior. They're all
  // subtly different.
  if (destOuterWidth !== srcOuterWidth) {
    dest.css({
      'min-width': srcOuterWidth + (srcOuterWidth - destOuterWidth)
    });
  }

  // They damn better well be equal now.
  assert(dest.outerWidth() === srcOuterWidth);
}

/**
 * Given two table cells, make dest the same outer height as src.
 */
function matchCellHeight(src: JQuery, dest: JQuery) {
  var srcOuterHeight: number = src.outerHeight(),
    destOuterHeight: number;

  dest.css({
    'height': srcOuterHeight
  });

  // After applying that change, how tall is the destination cell?
  destOuterHeight = dest.outerHeight();

  // Box model, how does it work?!
  // The truth is, it doesn't matter. Try something, see if we come up short
  // (or tall), and then adjust to the current browser's behavior. They're all
  // subtly different.
  if (destOuterHeight !== srcOuterHeight) {
    dest.css({
      'height': srcOuterHeight + (srcOuterHeight - destOuterHeight)
    });
  }

  // They damn better well be equal now.
  assert(dest.outerHeight() === srcOuterHeight);
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
  /**
   * Get the coordinates of the item.
   */
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

    for (i = 0; i < data.output.length; i++) {
      // Grab each item, create two-way links.
      item = <OutputItem> graph.getItem(data.output[i]);
      assert(item.getType() === DDType.OUTPUT, "Input dependents must be outputs.");
      // Output -> input
      item.addDependency(this);
      // Input -> output
      this.dependents.push({
        noerr: data.output[i].noerr,
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
   * this.data[worksheet][row][col]
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
    var row = wsData[coord.y];
    if (!row) {
      var width = wsData.length, i: number;
      for (i = width; i <= coord.y; i++) {
        wsData[i] = [];
      }
      row = wsData[coord.y] = [];
    }

    if (row.length < coord.x) {
      // Pad with empty entries.
      this.data[coord.worksheet][coord.y] = row = row.concat(new Array(coord.x - row.length + 1));
    }
    row[coord.x] = item;
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
  public getRow(ws: string, row: number): DDItem[] {
    var wsData = this.getWs(ws),
      rowData = wsData[row];
    if (typeof rowData === 'undefined') {
      throw new Error('Invalid row: ' + ws + ", " + row);
    }
    return rowData;
  }

  /**
   * Retrieves the given coordinate from the spreadsheet, or throws an error
   * if not found.
   */
  public getItem(coord: SpreadsheetCoordinate): DDItem {
    var row = this.getRow(coord.worksheet, coord.y), item = row[coord.x];
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
  private worksheetClassID: string;
  private width: number;
  private height: number;
  // Keeps track of cells that have changed since this spreadsheet was last displayed.
  private changedCells: SpreadsheetCoordinate[] = [];
  
  /* Page elements */
  
  /**
   * The ccTab <div> that encapsulates the entire worksheet.
   */
  private tableDiv: JQuery;
  /**
   * The column header <table>.
   */
  private colHeader: JQuery;
  /**
   * The row header <table>.
   */
  private rowHeader: JQuery;
  /**
   * The worksheet's tab <li>.
   */
  private tab: JQuery;
  /**
   * The worksheet's tab <a>.
   */
  private tabAnchor: JQuery;
  /**
   * The body of the spreadsheet, containing all of the data.
   */
  private body: JQuery;

  /**
   * @param data The data to display in the worksheet.
   * @param question The CheckCellQuestion that this worksheet belongs to.
   */
  constructor(private question: CheckCellQuestion, private name: string, private data: DDItem[][]) {
    this.width = this.calculateWidth();
    this.height = this.calculateHeight();
    this.worksheetClassID = this.question.getClassID() + 'Worksheet' + nextId();
    this.tableDiv = this.constructSkeleton();
    
    // Scroll row/col headers when body scrolls.
    this.tableDiv.find('.ccWorksheetBody-wrapper').scroll(function (e) {
      $('.ccRowHeader-wrapper').scrollTop($(this).scrollTop());
      $('.ccColHeader-wrapper').scrollLeft($(this).scrollLeft());
    });
  }
  
  public getClassID(): string {
    return this.worksheetClassID; 
  }
  
  public getTab(): JQuery {
    return this.tab; 
  }
  
  private calculateHeight(): number {
    return this.data.length;
  }
  
  private calculateWidth(): number {
    var i: number, maxLength: number = 0;
    for (i = 0; i < this.data.length; i++) {
      if (this.data[i].length > maxLength) {
        maxLength = this.data[i].length;  
      }
    }
    return maxLength;
  }
  
  private constructSkeleton(): JQuery {
    this.tabAnchor = $('<a>')
        .attr('href', '#' + this.getClassID())
        .text(this.name)
        .on('click', () => {
          this.updateHeaderCellSizes();
        });;
    this.tab = $('<li>')
      .append(this.tabAnchor);
    this.colHeader = this.constructColHeader();
    this.rowHeader = this.constructRowHeader();
    return $('<div class="ccTab ' + this.getClassID() + '" id="' + this.getClassID() + '">')
      .append($('<div class="ccTabContents">')
        .append($('<div class="ccColHeader-wrapper">')
          .append(this.colHeader)
        )
        .append($('<div class="ccBottom-wrapper">')
          .append($('<div class="ccRowHeader-wrapper">')
            .append(this.rowHeader)
          )
          .append($('<div class="ccWorksheetBody-wrapper">'))
       )
     ); 
  }
  
  public fillInSkeleton(): void {
    var bodyWrapper: JQuery = this.tableDiv.find('.ccWorksheetBody-wrapper'),
      i: number, j: number, row: DDItem[],
      rowHeaderHeader: JQuery = this.colHeader.find('.ccRowHeaderHeader');
    this.body = this.constructBody(); 
    bodyWrapper.append(this.body);
    // Insert into page.
    this.tableDiv.find('.ccBottom-wrapper').append(bodyWrapper);
    // Update rowHeaderHeader's width.
    matchCellWidth(this.rowHeader, rowHeaderHeader);
    // Update all cells, now that they are in the page.
    for (i = 0; i < this.data.length; i++) {
      row = this.data[i];
      for (j = 0; j < row.length; j++) {
        if (typeof row[j] !== 'undefined') {
          row[j].fireEvent('changed', row[j]);
        }
      }
    }
  }
  
  public getName(): string { return this.name; }

  public getDiv(): JQuery {
    return this.tableDiv;
  }

  public toggleHighlighting(enable: boolean) {
    if (enable !== this.tabAnchor.hasClass('ccTabChanged')) {
      if (enable) {
        this.tabAnchor.text(this.name + '*').addClass('ccTabChanged');
      } else {
        this.tabAnchor.text(this.name).removeClass('ccTabChanged');
      }
    }
  }

  private constructBlankCell(coords: SpreadsheetCoordinate): JQuery {
    var colClass: string = getColClass(this.getClassID(), coords.x),
      rowClass: string = getRowClass(this.getClassID(), coords.y);
    return $('<td class="' + colClass + ' ' + rowClass + '">');
  }

  private constructCell(data: DDItem): JQuery {
    var cell: JQuery = this.constructBlankCell(data.getCoords());
    if (data.getType() === DDType.INPUT) {
      if (!(<InputItem>data).isContext()) {
        cell.addClass('ccInput');
      }
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
            cell.removeClass('ccInputDisabledDraggable')
              .addClass('ccInputDraggable')
              .draggable({
                // appendTo needed for IE8:
                // http://stackoverflow.com/questions/14603785/jquery-ui-draggable-custom-helper-doesnt-work-correctly-in-ie7
                appendTo: 'body',
                cursor: 'move',
                revert: 'invalid',
                helper: () => {
                  return $('<li class="ccSortableListItem">' + coords2string(data.getCoords()) + ': ' + input.getErrorValue() + '</li>').data("DDItem", data);
                },
                scroll: true
              });
          } else {
            cell.removeClass('ccInputDraggable')
              .addClass('ccInputDisabledDraggable')
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

    var coords = data.getCoords();
    // Change events can be triggered by the global spreadsheet, *or* by the
    // above click handler.
    data.addEventListener('changed', (data: DDItem) => {
      // Update displayed value.
      cell.text(data.getValue());
      // Update headers.
      this.changedCells.push(data.getCoords());
      this.updateHeaderCellSizes();
      
      if (data.getType() === DDType.INPUT) {
        if (data.isValueErroneous()) {
          // Add the 'erroneous' style.
          cell.addClass('ccInputCellError');
        } else {
          // Remove the 'erroneous' style.
          cell.removeClass('ccInputCellError');
        }
      } else {
        // Check if we are in NO_ERROR or ALL_BUT_ONE_ERROR.
        if (this.question.getStatus() !== SpreadsheetStatus.ALL_ERRORS) {
          cell.addClass('ccChangedOutputCell');
          this.question.addToChangeList(<OutputItem>data);
        } else {
          cell.removeClass('ccChangedOutputCell');
          this.question.removeFromChangeList(<OutputItem>data);
        }
      }

      // Highlight our tab if this element is part of a single disabled
      // error.
      this.toggleHighlighting((!data.isValueErroneous()) && this.question.getStatus() === SpreadsheetStatus.ALL_BUT_ONE_ERROR);
    });
    return cell;
  }

  /**
   * Constructs a data row of the table.
   */
  private constructRow(row: DDItem[], rowId: number): JQuery {
    var tr: JQuery = $('<tr>'), i: number, item: DDItem, newCell: JQuery;

    // XXX: Excel is 1-indexed. Ignore the 0th cell.
    for (i = 1; i < this.width; i++) {
      item = row[i];
      if (typeof item === 'undefined') {
        newCell = this.constructBlankCell({worksheet: this.name, y: rowId, x: i});
      } else {
        newCell = this.constructCell(item);
      }
      tr.append(newCell);
    }
    return tr;
  }

  private constructRowHeader(): JQuery {
    var table: JQuery = $('<table class="ccRowHeader">'), i: number, tr: JQuery;
    // XXX: Excel is 1-indexed.
    for (i = 1; i < this.height; i++) {
      tr = $('<tr>');
      tr.append($('<td>')
        .text(i)
      );
      table.append(tr);
    }
    // Dummy row; accounts for extra scrollbar space.
    table.append($('<tr><td></td></tr>'));
    return table;
  }

  private constructColHeader(): JQuery {
    var table: JQuery = $('<table class="ccColHeader">'), i: number, tr = $('<tr>');
    // Construct header.
    tr.append($('<th class="ccRowHeaderHeader">'));
    // XXX: Excel is 1-indexed.
    for (i = 1; i < this.width; i++) {
      tr.append($('<th>' + getExcelColumn(i) + '</th>'));
    }
    // Dummy column; accounts for extra scrollbar space.
    tr.append($('<th></th>'));
    return table.append(tr);
  }
  
  /**
   * Constructs the <table> and its header.
   */
  private constructBody(): JQuery {
    var table: JQuery,
      i: number, tr: JQuery = $('<tr>');

    table = $('<table class="ccWorksheetBody">');
    // XXX: Excel is 1-indexed.
    for (i = 1; i < this.height; i++) {
      if (i < this.data.length) {
        table.append(this.constructRow(this.data[i], i));
      } else {
        table.append(this.constructRow([], i));
      }
    }
    return table;
  }

  /**
   * Find the cell element in the worksheet corresponding to the given
   * (col, row) coordinate. col = 0 and row = 0 searches the header row.
   */
  private findCell(col: number, row: number): JQuery {
    var rv: JQuery, rowElement: JQuery;
    if (row === 0) {
      // Handles (0, 0) edge case.
      rv = this.colHeader.find('tr th:nth-child(' + (col + 1) + ')');
    } else if (col === 0) {
      // Row header search. 0-indexed
      rv = this.rowHeader.find('tr:nth-child(' + row + ')').find(':nth-child(1)');
    } else {
      // General cell search.
      // Get row.
      rowElement = this.body.find('tr:nth-child(' + row + ')');
      assert(rowElement.length === 1);
      // Get col in row.
      rv = rowElement.find('td:nth-child(' + col + ')');
    }
    assert(rv.length === 1);
    return rv;
  }
  
  private updateColHeaderCellWidth(col: number) {
    matchCellWidth(this.findCell(col, 1), this.findCell(col, 0));
    // Due to fun border size issues, update adjacent columns. :|
    if (col < this.width - 1)
      matchCellWidth(this.findCell(col + 1, 1), this.findCell(col + 1, 0));
    if (col > 1)
      matchCellWidth(this.findCell(col - 1, 1), this.findCell(col - 1, 0));
  }
  
  private updateRowHeaderCellHeight(row: number) {
    matchCellHeight(this.findCell(1, row).parent(), this.findCell(0, row));
    // Due to fun border size issues, update adjacent rows. :|
    if (row < this.height - 1)
      matchCellHeight(this.findCell(1, row + 1), this.findCell(0, row + 1));
    if (row > 1)
      matchCellHeight(this.findCell(1, row - 1), this.findCell(0, row - 1));
  }
  
  /**
   * Is this the active worksheet tab?
   */
  private isActive(): boolean {
    return this.tableDiv.is(':visible')
  }
  
  
  private headerUpdatePending: boolean = false;
  /**
   * Updates all header cells (row and column) to be the appropriate height.
   */
  private updateHeaderCellSizes() {
    if (this.headerUpdatePending) return;
    // Need to wait until page is redrawn before widths/isActive updates.
    this.headerUpdatePending = true;
    setTimeout(() => {
      this.headerUpdatePending = false;
      // Ignore if we aren't active.
      if (this.isActive()) {
        var cell: SpreadsheetCoordinate, cols: number[] = [], rows: number[] = [],
          i: number;
        while(this.changedCells.length > 0) {
          cell = this.changedCells.pop();
          if (cols.indexOf(cell.x) === -1) cols.push(cell.x);
          if (rows.indexOf(cell.y) === -1) rows.push(cell.y);
        }
        for (i = 0; i < cols.length; i++) {
          this.updateColHeaderCellWidth(cols[i]); 
        }
        for (i = 0; i < rows.length; i++) {
          this.updateRowHeaderCellHeight(rows[i]); 
        }
      }
    }, 0);
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
 */
class CheckCellQuestion {
  private graph: DataDependencyGraph;
  private status: SpreadsheetStatus = SpreadsheetStatus.ALL_ERRORS;
  private tables: { [ws: string]: WorksheetTable } = {};
  private disabledError: InputItem = null;
  private questionClassID: string = 'ccQuestion' + nextId();
  
  /* Question elements */
  
  /**
   * The div that the question is situated in.
   */
  private questionDiv: JQuery;
  /**
   * The <button> that toggles all errors on/off.
   */
  private toggleButton: JQuery;
  /**
   * The changed items <ul> list.
   */
  private changedList: JQuery;
  /**
   * The ranked items <ul> list.
   */
  private rankedList: JQuery;
  /**
   * The unimportant items <ul> list.
   */
  private unimportantList: JQuery;
  /**
   * The tab <ul> list.
   */
  private tabList: JQuery;

  /**
   * @param data The JSON object with the question information.
   * @param divId The ID of the div where the question should be injected.
   */
  constructor(private data: QuestionInfo, parentDiv: JQuery) {
    this.graph = new DataDependencyGraph(data);
    this.questionDiv = $('<div class="ccQuestionDiv ' + this.getClassID() + '" >');
    this.questionDiv = this.constructSkeleton();
    parentDiv.append(this.questionDiv);
    setTimeout(() => {
      this.fillInSkeleton();
    }, 0);
  }
  
  /**
   * Constructs a skeleton of the question div.
   */
  private constructSkeleton(): JQuery {
    var changedListDiv: JQuery = this.constructChangedList();
    this.changedList = changedListDiv.find('ul');
    this.toggleButton = this.constructToggleButton();
    this.tabList = $('<ul class="ccTabList">');
    return $('<div class="ccQuestion ' + this.getClassID() + '">')
      .append($('<div class="ccTabs">').append(this.tabList))
      .append(changedListDiv)
      .append(this.constructInputLists())
      .append(this.toggleButton);
  }
  
  /**
   * Fills in the question div skeleton.
   */
  private fillInSkeleton(): void {
    var graphData = this.graph.getData(), i: number, ws: string, wsTable: WorksheetTable,
      tabsDiv: JQuery = this.questionDiv.find('.ccTabs');
    for (ws in graphData) {
      if (graphData.hasOwnProperty(ws)) {
        // Create worksheet.
        // NOTE: WST wants the anchor, not the list element.
        wsTable = new WorksheetTable(this, ws, graphData[ws]);
        this.tables[ws] = wsTable;
        // Append to page.
        this.tabList.append(wsTable.getTab());
        tabsDiv.append(wsTable.getDiv());
        // Fill it in.
        wsTable.fillInSkeleton();
      }
    } 
    this.questionDiv.tabs();
  }
  
  private constructSortableList(title: string, commonClass: string): JQuery {
    var self: CheckCellQuestion = this;
    return $('<div>')
      .append('<h4>' + title + '</h4>')
      .append($('<ul class="' + commonClass + '">')
        .sortable({
          revert: 'false',
          connectWith: '.' + commonClass,
          placeholder: 'ccPlaceholderItem'
        })
        .droppable({
          tolerance: 'pointer',
          accept: () => { return true; },
          drop: function (e, ui) {
            // Only append if the item is a child element of the tabs of the question div.
            var draggable: JQuery = $(ui.draggable);
            if (draggable.closest('.ccTab').length > 0 && draggable.closest('.' + self.getClassID()).length > 0) {
              var item: InputItem = ui.helper.data('DDItem'),
                helper: JQuery = ui.helper;
              $(this).append($('<li>').text(helper.text()).data('DDItem', item));
              // Wait one turn for jQuery UI to do it's thing before we disable
              // dragging. Otherwise, strange things happen.
              setTimeout(() => { item.setDraggable(false); }, 0);
            }
          }
        })
      );
  }
  
  private constructInputLists(): JQuery {
    var commonID: string = this.getClassID() + "InputLists",
      rankedListDiv: JQuery = this.constructSortableList('Ranked Inputs', commonID),
      unimportantListDiv: JQuery = this.constructSortableList('Unimportant Inputs', commonID);
    this.rankedList = rankedListDiv.find('ul');
    this.unimportantList = unimportantListDiv.find('ul');
    return $('<div class="ccListsDiv">')
      .append(rankedListDiv)
      .append(unimportantListDiv); 
  }
  
  private constructToggleButton(): JQuery {
    return $('<button>')
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
  }
  
  private constructChangedList(): JQuery {
    return $('<div class="ccChangedList"><h5>Changed Outputs</h5><ul></ul></div>');
  }
  
  public getClassID(): string {
    return this.questionClassID; 
  }
  
  public addToChangeList(output: OutputItem): void {
    // Race condition :(
    if (this.changedList == null) return;
    var coords = coords2string(output.getCoords());
    if (this.changedList.find("li:contains('" + coords + "')").length === 0) {
      this.changedList.append($('<li>').text(coords));
    }
  }

  public removeFromChangeList(output: OutputItem): void {
    var coords = coords2string(output.getCoords());

    this.changedList.find("li:contains('" + coords + "')").remove();
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
    var unimportantList = this.unimportantList.children();
    for (i = 0; i < unimportantList.length; i++) {
      // XXX: Hack cuz list values are "coords: value".
      coords = $(unimportantList[i]).text().split(':')[0];
      item = coords2item[coords];
      delete coords2item[coords];
      assert(typeof item !== 'undefined');
      rv.unimportant.push(item.getCoords());
    }

    // Find each item in the rank list in the hash.
    var rankList = this.rankedList.children();
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

// Onload code. Fixes draggable elements in Firefox. Without this, draggable
// elements will appear severely offset from the cursor in that browser.
// Also Opera.
if ((<any>$).browser.mozilla || (<any>$).browser.opera) {
  $(document).ready(() => {
    $('body').css('position', 'relative');
  });
}
