/**
 * BookShelfクラスのクラス変数情報<br />
 * 
 * @typedef {Object} BookShelfSheetParam
 * 
 * @prop {?number}        bookListAreaStartColumn     ブックリスト範囲の開始列番号
 * @prop {?RangeAddress}  rootFolderUrlCellAddress    ルートフォルダURLのセル情報
 * @prop {?RangeAddress}  templateBookUrlCellAddress  テンプレートBookUrlのセル情報
 * 
 */

/**
 * BookShelfのテンプレートファイル名の変換パラメータ情報<br />
 * 
 * @typedef {Object} renameTemplateBookNameParam
 * 
 * @prop {?string}  yyyy    埋め込んでいる【yyyy】（年）の変更後の値
 * @prop {?string}  mm      埋め込んでいる【mm】（月）の変更後の値
 * @prop {?string}  dd      埋め込んでいる【dd】（日）の変更後の値
 * @prop {?string}  suffix  ファイル末尾に設定する値
 * 
 */

/**
 * BookShelfSheetを管理するテンプレートClass
 */
class BookShelfSheet {
  /**
   * BookShelfインスタンスを作成する
   * 
   * @param {SheetControllerParam} param  初期化パラメータ
   */
  constructor(param) {
    /** 
     * BookShelf用のSheetController
     * @type {SheetController}
     * @private
     */
    const keyColumn = 2;  // 2列目をキー列（入力必須列）とする
    this._sheetController = new SheetController({
      sheetName: "ブック一覧",
      headerRow: 13,
      keyColumn: keyColumn
    });

    /** 
     * Bookの追加時の開始列番号
     * @type {number}
     * @private
     */
    this._bookListAreaStartColumn = param.bookListAreaStartColumn || keyColumn + 1;

    /** 
     * ルートドライブURL保持のセル情報
     * @type {RangeAddress}
     * @private
     */
    this.rootFolderUrlCellAddress = param.rootFolderUrlCellAddress || {startRow: 3, startColumn: 2};

    /** 
     * テンプレートBook保持のセル情報
     * @type {RangeAddress}
     * @private
     */
    this._templateBookUrlCellAddress = param.templateBookUrlCellAddress || {startRow: 6, startColumn: 2};
  }

  /**
   * BookShelfシートのSheetControllerを取得する
   * 
   * @return {SheetController} BookShelfシートのSheetController
   * 
   */
  getBookShelfSheetController() {
    return this._sheetController;
  }

  /** 
   * Bookの保管場所となるルートDriveのFolderオブジェクトを取得する
   * 
   * @return {DriveApp.Folder} BookShelfシートで管理するBookの保管場所となるルートフォルダオブジェクト
   * 
   */
  getRootFolder() {
    const rootFolderUrl = this._sheetController.getRange(this.rootFolderUrlCellAddress).getValue();
    const rootFolderId = BookController.extractFolderIdFromUrl(rootFolderUrl);

    return DriveApp.getFolderById(rootFolderId);
  }

  /** 
   * BookShelfで管理するテンプレートBookのBookControllerを取得する
   * 
   * @return {BookController} BookShelfシートで管理するテンプレートのBookController
   * 
   */
  getTemplateBookController() {
    const templateBookUrl = this._sheetController.getRange(this._templateBookUrlCellAddress).getValue();

    return new BookController({
      bookId: BookController.extractBookIdFromUrl(templateBookUrl)
    });
  }

  /** 
   * BookShelfで管理するBookのテンプレートからコピーを作成し、そのコピーのBookControllerを取得する
   * 
   * @param {?string} [createFolderUrl=null]  テンプレートをコピーしたファイルの保管フォルダURL
   * 
   * @return {BookController} BookShelfシートで管理するテンプレートのBookController
   * 
   */
  createBookControllerFromTemplate(createBookName=null, createFolderUrl=null) {
    /**
     * @type {BookControllerParam}
     */
    const param = {};
    
    if(createBookName) {
      param.bookName = createBookName;
    }

    if(createFolderUrl) {
      param.folderId = BookController.extractFolderIdFromUrl(createFolderUrl);
    } else {
      // URLの指定がない場合はルートフォルダに作成する
      param.folderId = this.getRootFolder().getId();
    }

    return BookController.createBookFromAnotherBook(this.getTemplateBookController().getBook().getId(), param);
  }

  /** 
   * Bookの保管場所となるルートDriveに子フォルダを作成する
   * 
   * @param {!string} createFolderName  作成するフォルダ名
   * 
   * @return {DriveApp.Folder} 作成したフォルダオブジェクト
   * 
   */
  createFolderInRootFolder(createFolderName) {
    if(!createFolderName) {
      throw new Error("指定されたフォルダ名が空文字です。有効なフォルダ名を指定してください");
    }

    const rootFolder = this.getRootFolder();
    const childFolders = rootFolder.getFoldersByName(createFolderName);
    let createFolder;

    if(childFolders.hasNext()) {
      // 指定フォルダ名のフォルダが存在すればそれを返す
      createFolder = childFolders.next();
    } else {
      createFolder = rootFolder.createFolder(createFolderName);
    }

    return createFolder;
  }

  /** 
   * チェックされている行番号を特定する
   * 
   * @return {Array<number>} BookShelfシートのチェックボックスがチェックされている行番号
   * 
   */
  getCheckedDataAreaRowArray() {
    const checkBoxColumnArray = this._sheetController.getColumnRange(
      BookShelfSheet.CHECKBOX_COLUMN
    ).getValues().flat();
    const {dataStartRow} = this._sheetController.getSheetControllerParam();

    const checkedRowArray = [];
    checkBoxColumnArray.forEach((row, index) => {
      if(row) {
        checkedRowArray.push(dataStartRow + index);
      }
    });

    return checkedRowArray;
  }

  /**
   * BookShelfシートのデータエリアに1行を追加する
   * 
   * @param {?string} [addPosition=null]  追加する位置（指定がない場合は先頭行）
   * 
   * @return {number} 追加した行番号
   * 
   */
  addRowInDataArea(addPosition=null) {
    const {dataStartRow, lastRow} = this._sheetController.getSheetControllerParam();
    let addRow;
    if(addPosition === BookShelfSheet.ADD_POSITION_LAST) {
      addRow = lastRow + 1;
    } else {
      addRow = dataStartRow;
    }

    const newAdditionalRangeRow = this._sheetController.getRowRange(addRow).insertCells(SpreadsheetApp.Dimension.ROWS).getRow();

    // 追加列にチェックボックスを追加する
    this._sheetController.getRange({
      startRow: newAdditionalRangeRow,
      startColumn: BookShelfSheet.CHECKBOX_COLUMN
    }).insertCheckboxes();

    return newAdditionalRangeRow;
  }

  /**
   * BookShelfシートのブックリストエリアに1列を追加する
   * 
   * @param {!string} addColumnName       追加する列名
   * @param {?string} [addPosition=null]  追加する位置（指定がない場合はブックリストエリアの先頭列）
   * @param {?string} [columnUrl=null]    追加列の遷移先URL（指定がない場合はリンクの設定を行わない）
   * 
   * @return {number} 追加した列番号
   * 
   */
  addColumnInBookArea(addColumnName, addPosition=null, columnUrl=null) {
    const {lastColumn, headerRow} = this._sheetController.getSheetControllerParam();
    let addColumn;
    if(addPosition === BookShelfSheet.ADD_POSITION_LAST) {
      addColumn = lastColumn + 1;
    } else {
      addColumn = this._bookListAreaStartColumn;
    }

    const targetColumnRange = this._sheetController.getColumnRange(addColumn, 1, false, true);
    const newAdditionalRangeColumn = targetColumnRange.insertCells(SpreadsheetApp.Dimension).getColumn();

    // 追加列のヘッダーに列名をセットする

    // URLチェックがNGの場合、リンク設定は行わない
    let setValueStr = addColumnName;
    const urlRegExp = new RegExp(/^https:\/\//);
    
    if(urlRegExp.test(columnUrl)) {
      setValueStr = `=HYPERLINK("${columnUrl}", "${addColumnName}")`;
    }

    this._sheetController.getRange({
      startRow: headerRow,
      startColumn: newAdditionalRangeColumn
    }).setValue(setValueStr);

    return newAdditionalRangeColumn;
  }

  /** 
   * 新しく入力エリアを追加する位置（先頭）
   * 
   * @return {string} 入力エリアの追加位置（先頭）
   * 
   */
  static get ADD_POSITION_FIRST() {
    return "add_position_first";
  }

  /** 
   * 新しく入力エリアを追加する位置（最終）
   * 
   * @return {string} 入力エリアの追加位置（最終）
   * 
   */
  static get ADD_POSITION_LAST() {
    return "add_position_last";
  }

  /** 
   * チェックボックス列の列番号
   * 
   * @return {number} チェックボックス列の列番号
   * 
   */
  static get CHECKBOX_COLUMN() {
    return 1;
  }

  /** 
   * 実行日から指定日日付をさかのぼった日付のテンプレートファイル名のリネームパラメータを作成する
   * 
   * @param {?number} [agoYearNum=0]  さかのぼる年数
   * @param {?number} [agoMonthNum=0] さかのぼる月数
   * @param {?number} [agoDayNum=0]   さかのぼる日数
   * 
   * @return {renameTemplateBookNameParam} リネーム後ファイル名
   * 
   */
  static createRenameParamFromExecDate(agoYearNum=0, agoMonthNum=0, agoDayNum=0) {
    const now = new Date();
    const targetDate = new Date(
      now.getFullYear() - agoYearNum,
      now.getMonth() - agoMonthNum,
      now.getDate() - agoDayNum
    );

    return {
      yyyy: String(targetDate.getFullYear()),
      mm: `0${targetDate.getMonth() + 1}`.slice(-2),
      dd: `0${targetDate.getDate()}`.slice(-2),
    };
  }

  /** 
   * テンプレートBookのファイル名に埋め込んでいたパラメータを変換し、リネームしたファイル名を取得する
   * 
   * @param {string}                      templateFileName  変更前のテンプレートファイル名
   * @param {renameTemplateBookNameParam} renameParam       変更後パラメータ
   * 
   * @return {string} リネーム後ファイル名
   * 
   */
  static getRenameTemplateBookName(templateFileName, renameParam) {
    let yyyy;
    let mm;
    let dd;
    let suffix;

    if(renameParam) {
      yyyy = renameParam.yyyy;
      mm = renameParam.mm;
      dd = renameParam.dd;
      suffix = renameParam.suffix;
    }

    // ファイル名の【テンプレート】をなくす
    let afterRenameName = templateFileName.replace("【テンプレート】", "");

    if(yyyy) {
      afterRenameName = afterRenameName.replace("【yyyy】", yyyy);
    }

    if(mm) {
      afterRenameName = afterRenameName.replace("【mm】", mm);
    }

    if(dd) {
      afterRenameName = afterRenameName.replace("【dd】", dd);
    }

    if(suffix) {
      afterRenameName = `${afterRenameName}_${suffix}`;
    }

    return afterRenameName;
  }
}
