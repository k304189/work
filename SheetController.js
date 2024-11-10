/**
 * SheetControllerクラスで操作するRangeの範囲情報の型
 * @typedef {Object} RangeAddress
 * 
 * @prop {?string}  rangeName   範囲名
 * @prop {!number}  startRow    開始行
 * @prop {!number}  startColumn 開始列
 * @prop {?number}  rowNum      対象行数
 * @prop {?number}  columnNum   対象列数
 */

/**
 * SheetControllerクラスの初期化の必要情報の型
 * @typedef {Object} SheetControllerParam
 * 
 * @prop {?string}  bookId            対象ブックID
 * @prop {!string}  sheetName         操作対象シート名
 * @prop {?number}  headerRow         ヘッダー行番号
 * @prop {?number}  startColumn       開始列番号
 * @prop {?number}  keyColumn         キー（入力必須）列番号
 */

/**
 * ブック内のシートを操作する処理をまとめたテンプレートクラス<br />
 */
class SheetController {
  /**
   * SheetControllerインスタンスを作成する
   * 
   * @param {SheetControllerParam} param  初期化パラメータ
   */
  constructor(param) {
    this.setSheetControllerParam(param);
  }

  /**
   * SheetControllerのクラス変数を更新する関数<br />
   * <ul>
   *  <li>パラメータに指定がない項目は更新せず、現状の設定のままとする</li>
   *  <li>また最初の初期化に関しては、初期値を設定する</li>
   * </ul>
   * @param {SheetControllerParam} param 更新する設定値
   * 
   */
  setSheetControllerParam(param) {
    let changeBookId = false;
    // ブックインスタンス
    if(param.bookId) {
      try {
        /** 
         * 操作シートを持っているブックインスタンス
         * @type {SpreadsheetApp.Spreadsheet}
         * @private
         */
        this._book = SpreadsheetApp.openById(param.bookId);
        changeBookId = true;
      } catch(e) {
        console.log(e.stack);
        throw new Error("指定されたブックを開くことができませんでした");
      }
    } else if(!this._book) {  // クラス変数にブックインスタンスが存在しない場合は現状のブックをセットする
      this._book = SpreadsheetApp.getActiveSpreadsheet();
      changeBookId = true;
    }

    // シートインスタンス
    let changeSheetName = false;
    // シート名の変更もしくは、参照ブックが変更されたときに再取得する
    if(param.sheetName || changeBookId) {
      /** 
       * 操作シート名
       * @type {string}
       * @private
       */
      this._sheetName = param.sheetName || this._sheetName;
      const hasSheetNames = this._book.getSheets().map(sheet => sheet.getSheetName());
      
      // 対象シート名のシートが存在しない場合は新しく作成する
      if(hasSheetNames.includes(this._sheetName)) {
        /** 
         * 操作シートのインスタンス
         * @type {SpreadsheetApp.Sheet}
         * @private
         */
        this._sheet = this._book.getSheetByName(this._sheetName);
      } else {
        this._sheet = this._book.insertSheet(this._sheetName);
      }
      changeSheetName = true;
    }
    
    // ヘッダー行番号
    let changeHeaderRow = false;
    if(param.headerRow) {
      /** 
       * 操作シートのヘッダー行
       * @type {number}
       * @private
       */
      this._headerRow = param.headerRow;
      changeHeaderRow = true;
    } else if(!this._headerRow) { // 初めての初期化かつ指定がない場合は1行目をヘッダー行とする
      this._headerRow = 1;
      changeHeaderRow = true;
    }

    // データ開始行番号
    if(changeHeaderRow) {
      /** 
       * 操作シートのデータ開始行
       * @type {number}
       * @private
       */
      this._dataStartRow = this._headerRow + 1;
    }

    // 開始列番号
    let changeStartColumn = false;
    if(param.startColumn) {
      /** 
       * 操作シートのデータエリア開始列
       * @type {number}
       * @private
       */
      this._startColumn = param.startColumn;
      changeStartColumn = true;
    } else if(!this._startColumn) { // 初めての初期化かつ指定がない場合は1行目をヘッダー行とする
      this._startColumn = 1;
      changeStartColumn = true;
    }

    // キー（入力必須）列番号
    if(param.keyColumn) {
      /** 
       * 操作シートのキー（入力必須）列番号
       * @type {number}
       * @private
       */
      this._keyColumn = param.keyColumn
    } else if(changeStartColumn) {
      // キー列の指定がなく、開始列が変わったときは開始列に合わせる
      this._keyColumn = this._startColumn;
    }

    // ヘッダー行番号・開始列番号が変更された場合、関連するクラス変数を変更する
    if(changeBookId || changeSheetName || changeHeaderRow || changeStartColumn) {
      this.reflashLastRow();
      this.reflashLastColumn();
    }
  }

  /**
   * 最終行に関連するクラス変数を更新する
   */
  reflashLastRow() {
    /** 
     * 操作シートの最終行番号
     * @type {number}
     * @private
     */
    this._lastRow = this.getLastRow();
    
    /** 
     * 操作シートのデータ行数
     * @type {number}
     * @private
     */
    this._rowNum = this._lastRow - this._dataStartRow + 1;
  }

  /**
   * 最終列に関連するクラス変数を更新する
   */
  reflashLastColumn() {
    /** 
     * 操作シートの最終列番号
     * @type {number}
     * @private
     */
    this._lastColumn = this.getLastColumn();
    
    /** 
     * 操作シートのデータエリア列数
     * @type {number}
     * @private
     */
    this._columnNum = this._lastColumn - this._startColumn + 1;
    
    /** 
     * 操作シートのヘッダーの列名の配列
     * @type {Array<string>}
     * @private
     */
    this._header = this.getHeaderRange().getValues().flat();
  }

  /**
   * 指定した範囲名のシートの最終行を取得する
   * 
   * @return {number} 対象範囲の最終行番号
   * 
   */
  getLastRow() {
    const baseCell = this._sheet.getRange(this._sheet.getMaxRows(), this._keyColumn);
    let lastRow = baseCell.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

    // 最終行に入力値があり、一番上のNextDataCellがヘッダー以前の行であれば、最終行まで値が入っているとする
    if(baseCell.getValue() && lastRow <= this._headerRow) {
      lastRow = baseCell.getRow();
    }
    return lastRow;
  }

  /**
   * 指定した範囲名のシートの最終列を取得する
   * 
   * @return {number} 対象範囲の最終列番号
   * 
   */
  getLastColumn() {
    const baseCell = this._sheet.getRange(this._headerRow, this._startColumn);
    return baseCell.getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
  }

  /**
   * クラスに紐づいているSheetインスタンスを取得する
   * 
   * @return {SpreadsheetApp.Sheet} クラスに紐づくSheetインスタンス
   * 
   */
  getSheet() {
    return this._sheet;
  }

  /**
   * シートから対象範囲を取得する<br />
   * rowNumとcolumnNumの指定がない場合は両方とも1として扱う（1セルを取得）
   * 
   * @param {RangeAddress} param  取得範囲の情報
   * 
   * @return {SpreadsheetApp.Range} 指定範囲のRangeインスタンス
   * 
   */
  getRange(param) {
    const startRow = param.startRow;
    const startColumn = param.startColumn;
    const rowNum = param.rowNum || 1;
    const columnNum = param.columnNum || 1;

    return this._sheet.getRange(startRow, startColumn, rowNum, columnNum);
  }

  /**
   * データエリアから指定した行の範囲を取得する<br />
   * allColumnにtrueを指定するとシート全列を対象とする
   * 
   * @param {number}  startRow          開始行番号
   * @param {number}  [rowNum=1]        取得行数
   * @param {boolean} [allColumn=false] 全列取得判定
   * 
   * @return {SpreadsheetApp.Range} 指定行のRangeインスタンス
   * 
   */
  getRowRange(startRow, rowNum=1, allColumn=false) {
    const startColumn = (allColumn) ? 1 : this._startColumn;
    const columnNum = (allColumn) ? this._sheet.getMaxColumns() : this._columnNum;
    const rowRange = this.getRange({
      startRow: startRow,
      startColumn: startColumn,
      rowNum: rowNum,
      columnNum: columnNum
    });

    return rowRange;
  }

  /**
   * データエリアから指定した列の範囲を取得する<br />
   * allRowにtrueを指定するとシート全行を対象とする
   * 
   * @param {number}  startColumn       開始列番号
   * @param {number}  [columnNum=1]     取得列数
   * @param {boolean} [allRow=false] 前列取得判定
   * 
   * @return {SpreadsheetApp.Range} 指定列のRangeインスタンス
   * 
   */
  getColumnRange(startColumn, columnNum=1, allRow=false) {
    const startRow = (allRow) ? 1 : this._dataStartRow;
    const rowNum = (allRow) ? this._sheet.getMaxRows() : this._rowNum;
    const columnRange = this.getRange({
      startRow: startRow,
      startColumn: startColumn,
      rowNum: rowNum,
      columnNum: columnNum
    });

    return columnRange;
  }

  /**
   * シートのヘッダー範囲を取得する
   * 
   * @return {SpreadsheetApp.Range} ヘッダーのRangeインスタンス
   * 
   */
  getHeaderRange() {
    const headerRange = this.getRange({
      startRow: this._headerRow,
      startColumn: this._startColumn,
      columnNum: this._columnNum
    });

    return headerRange;
  }

  /**
   * シートのヘッダーにデータを出力する
   * 
   * @param {Array<string|number|boolean>} headerArray 出力ヘッダー配列
   * 
   * @return {SpreadsheetApp.Range} 出力範囲のRangeインスタンス
   * 
   */
  outputHeader(headerArray) {
    const outputHeaderRange = this.getRange({
      startRow: this._headerRow,
      startColumn: this._startColumn,
      columnNum: headerArray.length
    });

    outputHeaderRange.setValues([headerArray]);

    this.reflashLastColumn();

    return outputHeaderRange;
  }

  /**
   * データエリアの範囲を取得する
   * 
   * @return {SpreadsheetApp.Range} データエリアのRangeインスタンス
   * 
   */
  getDataAreaRange() {
    const dataAreaRange = this.getRange({
      startRow: this._dataStartRow,
      startColumn: this._startColumn,
      rowNum: this._rowNum,
      columnNum: this._columnNum
    });

    return dataAreaRange;
  }

  /**
   * データエリアへデータを出力する
   * 
   * @param {Array<Array>} outputDataArray 出力データ配列
   * 
   * @return {SpreadsheetApp.Range} 出力範囲のRangeインスタンス
   * 
   */
  outputDataArea(outputDataArray) {
    const outputDataRange = this.getRange({
      startRow: this._dataStartRow,
      startColumn: this._startColumn,
      rowNum: outputDataArray.length,
      columnNum: outputDataArray[0].length
    });

    outputDataRange.setValues(outputDataArray);

    this.reflashLastRow();

    return outputDataRange;
  }

  /**
   * ヘッダー配列から指定の列名の配列添字を取得する
   * 
   * @param {string} columnName 添字を取得する列名
   * 
   * @return {number} 対象列名の添字
   * 
   */
  getHeaderArrayIndexByColumnName(columnName) {
    const idx = this._header.indexOf(columnName);
    if(idx === -1) {
      throw new Error("指定列名がヘッダーに存在しません");
    }
    return idx;
  }

  /**
   * ヘッダーから指定の列名のシート上の列番号を取得する
   * 
   * @param {string} columnName 添字を取得する列名
   * 
   * @return {number} 対象列名のシート状の列番号
   * 
   */
  getHeaderColumnByColumnName(columnName) {
    const headerArrayIdx = this.getHeaderArrayIndexByColumnName(columnName);
    return headerArrayIdx + this._startColumn;
  }

  /**
   * 自身のシートを別ブックにコピーする
   * 
   * @param {string} copyToBookId            コピー先のBookId
   * @param {string} [copyToSheetName=null]  コピー先のシート名（指定がない場合、シート名をそのまま使う）
   * 
   * @return {SheetController} コピーしたシートのコピー先ブックのSheetController
   * 
   */
  copySheetToAnotherBook(copyToBookId, copyToSheetName=null) {
    try {
      const sheetName = copyToSheetName || this._sheetName;
      const copyToBook = SpreadsheetApp.openById(copyToBookId);
      const copyToSheet = this._sheet.copyTo(copyToBook);
      copyToSheet.setName(sheetName);

      const copyToSheetController = Object.assign(Object.create(SheetController.prototype), this);
      copyToSheetController.setSheetControllerParam({
        bookId: copyToBookId,
        sheetName: sheetName
      });

      return copyToSheetController;
    } catch(e) {
      console.log(e.stack);
      throw new Error("別ブックへのSheetコピーが失敗しました");
    }
  }

  /**
   * 数字フォーマット：「書式なしテキスト」
   * @type {string}
   */
  static get NUMBER_FOMAT_NONE() {
    return "@";
  }

  /**
   * 「数値」「通貨」フォーマットをカスタマイズする
   * @private
   * 
   * @param {string}  defaultFormat         カスタマイズする基本フォーマット
   * @param {number}  [degitOfDecimal=0]    表示小数点桁数
   * @param {boolean} [whenMinusIsRed=true] マイナス時の赤字表示（デフォルト赤字）
   * 
   * @return {string} 数字フォーマット：「通貨」
   */
  static _createFormatNumber(defaultFormat, degitOfDecimal=0, whenMinusIsRed=true) {
    let decimalFormat = "";
    if(degitOfDecimal > 0) {
      decimalFormat = `.${"0".repeat(degitOfDecimal)}`;
    }

    let format = `${defaultFormat}${decimalFormat}`;

    if(whenMinusIsRed) {
      format = `${format}_);[Red]-${format}`;
    }
    return format;
  }

  /**
   * 数字フォーマット：「数値」のフォーマットをカスタマイズして取得する
   * 
   * @param {number}  [degitOfDecimal=0]    表示小数点桁数
   * @param {boolean} [whenMinusIsRed=true] マイナス時の赤字表示（デフォルト赤字）
   * 
   * @return {string} 数字フォーマット：「通貨」
   */
  static getNumberFormatNumber(degitOfDecimal=0, whenMinusIsRed=true) {
    return SheetController._createFormatNumber("#,##0", degitOfDecimal, whenMinusIsRed);
  }

  /**
   * 数字フォーマット：「通貨」のフォーマットをカスタマイズして取得する
   * 
   * @param {number}  [degitOfDecimal=0]    表示小数点桁数
   * @param {boolean} [whenMinusIsRed=true] マイナス時の赤字表示（デフォルト赤字）
   * 
   * @return {string} 数字フォーマット：「通貨」
   */
  static getNumberFormatCurrency(degitOfDecimal=0, whenMinusIsRed=true) {
    return SheetController._createFormatNumber("¥#,##0", degitOfDecimal, whenMinusIsRed);
  }

  /**
   * 数字フォーマット：「数値」のデフォルトフォーマット（カンマ区切り、小数点表示なし）
   * @type {string}
   */
  static get NUMBER_FOMAT_NUMBER_DEFAULT() {
    return SheetController.getNumberFormatNumber(0, false);
  }

  /**
   * 数字フォーマット：「通貨」のデフォルトフォーマット（カンマ区切り、小数点表示なし、マイナス赤字）
   * @type {string}
   */
  static get NUMBER_FOMAT_CURRENCY_DEFAULT() {
    return SheetController.getNumberFormatCurrency(0, false);
  }

  /**
   * 数字フォーマット：「パーセント」（小数第二位まで表示）
   * @type {string}
   */
  static get NUMBER_FOMAT_PERCENT() {
    return "0.00%";
  }
}
