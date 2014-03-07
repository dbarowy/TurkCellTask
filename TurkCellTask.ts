/**
 * JSON object that represents a single question.
 */
interface QuestionInfo {
  table: string[][];
}

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
 * Represents a single CheckCell task. Given a JSON object and a div id, it will
 * display the specified ranking question in the given div.
 */
class CheckCellQuestion {
  private width: number = 0;
  private table: HTMLTableElement;
  /**
   * @param data The JSON object with the question information.
   * @param divId The ID of the div where the question should be injected.
   */
  constructor(private data: QuestionInfo, divId: string) {
    var i: number;
    // Calculate width.
    for (i = 0; i < this.data.table.length; i++) {
      if (this.data.table[i].length > this.width) {
        this.width = this.data.table[i].length;
      }
    }

    // Construct the table.
    this.table = this.constructTable();
    for (i = 0; i < this.data.table.length; i++) {
      this.constructRow(this.table, this.data.table[i], i+1);
    }

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
    tr.appendChild(document.createElement('th'));
    for (i = 0; i < this.width; i++) {
      th = document.createElement('th');
      th.innerText = getExcelColumn(i);
      tr.appendChild(th);
    }
    table.appendChild(tr);
    return table;
  }
}


var sampleQuestion: QuestionInfo = {
  table: [
    ['100', '330', '284', '2856', '233'],
    ['3', '566', '3', '466', '32343']
  ]
};

window.onload = function () {
  var sampleTable = new CheckCellQuestion(sampleQuestion, 'sample');
};
