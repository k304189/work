/**
 * BookControllerクラスの初期化の必要情報の型
 * @typedef {Object} BookControllerParam
 * 
 * @prop {?string}  bookId    対象ブックID
 * @prop {?string}  folderId  ブックの配置フォルダID
 * @prop {?string}  bookName  ブック名
 * 
 */

/**
 * ブックを操作する処理をまとめたテンプレートクラス
 */
class BookController {
  /**
   * BookControllerインスタンスを作成する
   * 
   * @param {BookControllerParam} param  初期化パラメータ
   */
  constructor(param) {
    if(param.bookId) {
      try {
        /** 
         * 対象ブックインスタンス
         * @type {SpreadsheetApp.Spreadsheet}
         * @private
         */
        this._book = SpreadsheetApp.openById(param.bookId);
      } catch(e) {
        console.log(e.stack);
        throw new Error("指定されたブックを開くことができませんでした");
      }
    } else if(param.bookName) {
      this._book = SpreadsheetApp.create(param.bookName);
    } else {
      throw new Error("BookControllerのパラメータ指定がありません");
    }

    /** 
     * 対象ブックのファイルインスタンス
     * @type {DriveApp.File}
     * @private
     */
    this._bookFile = DriveApp.getFileById(this._book.getId());
    if(param.bookName) {
      this._bookFile.setName(param.bookName);
    }

    if(param.folderId) {  // folderIdがあれば、そのフォルダに移動させる
      try {
        /** 
         * 対象ブックインスタンス
         * @type {DriveApp.Folder}
         * @private
         */
        this._bookInFolder = DriveApp.getFolderById(param.folderId);
        this._bookFile.moveTo(this._bookInFolder);
      } catch(e) {
        console.log(e.stack);
        throw new Error("BookControllerの指定フォルダにアクセスできませんでした");
      }
    } else {
      const parentFolders = this._bookFile.getParents();
      if(parentFolders.hasNext()) {
        this._bookInFolder = parentFolders.next();
      }
    }
  }

  /** 
   * クラスに紐づくブックのSpreadsheetインスタンスを返す
   * 
   * @return {SpreadsheetApp.Spreadsheet} ブックのSpreadsheetインスタンス
   * 
   */
  getBook() {
    return this._book;
  }

  /** 
   * クラスに紐づくブックのFileインスタンスを返す
   * 
   * @return {DriveApp.File} ブックのFileインスタンス
   * 
   */
  getBookFile() {
    return this._bookFile;
  }

  /** 
   * クラスに紐づくブックが配置されているFolderインスタンスを返す
   * 
   * @return {DriveApp.Folder} ブックが配置されているFolderインスタンス
   * 
   */
  getBookInFolder() {
    return this._bookInFolder;
  }

  /** 
   * BookControllerに紐づくブックの配置フォルダを移動する
   * 
   * @param {string}  moveToFolderId  移動先フォルダID
   * 
   */
  moveFolder(moveToFolderId) {
    moveToFolder = DriveApp.getFolderById(moveToFolderId);
    this._bookFile.moveTo(moveToFolder);
    this._bookInFolder = moveToFolder;
  }

  /** 
   * ブックに同ドメインユーザーへの参照権限を付与する<br />
   * ファイル・フォルダに設定するために必要な列挙体は以下を参照
   * <ul>
   *   <li>[Enum Access]{@link https://developers.google.com/apps-script/reference/drive/access}</li>
   *   <li>[Enum Permission]{@link https://developers.google.com/apps-script/reference/drive/permission}</li>
   * </ul>
   * 
   */
  setSharingDomainView() {
    this._bookFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  }

  /**
   * 別ブックのシートを取得して、自分のブックに取り込む
   * 
   * @param {string}                fetchBookId           シート取得先ブックID
   * @param {string}                fetchSheetName        取得シート名
   * @param {SheetControllerParam}  sheetControllerParam  取り込みシートのSheetControllerパラメータ
   * 
   * @return {SheetController} 自ブックの取得したシートのSheetController
   * 
   */
  fetchSheetFromAnotherBook(fetchBookId, fetchSheetName, sheetControllerParam) {
    try {
      const fetchBook = SpreadsheetApp.openById(fetchBookId);
      const fetchSheet = fetchBook.getSheetByName(fetchSheetName);

      const createSheet = fetchSheet.copyTo(this._book);
      const sheetName = sheetControllerParam.sheetName || fetchSheetName;
      createSheet.setName(sheetName);

      // シート名の変換が必要か確認する
      sheetControllerParam.bookId = this._book.getId();
      sheetControllerParam.sheetName = sheetName;

      const controller = new SheetController(sheetControllerParam);
      return controller;
    } catch(e) {
      console.log(e.stack);
      throw new Error("他ブックからのシート取得が失敗しました");
    }
  }

  /** 
   * GoogleドライブのフォルダURLからフォルダIDを抽出する
   * 
   * @param {string} folderUrl  フォルダIDを抽出するGoogleドライブのフォルダURL
   * 
   * @return {string} フォルダID
   * 
   */
  static extractFolderIdFromUrl(folderUrl) {
    const urlRegExp = new RegExp("https://drive.google.com/drive/folders/([^/]+)");
    const matchResult = folderUrl.match(urlRegExp);
    let folderId;
    
    if(matchResult && matchResult.length >= 2) {
      folderId = matchResult[1];  // キャプチャグループは2つ目の要素に入っているので1を指定
    }
    return folderId;
  }

  /** 
   * GoogleドライブのブックURLからファイルIDを抽出する
   * 
   * @param {string} folderUrl  ファイルIDを抽出するGoogleドライブのブックURL
   * 
   * @return {string} ブックID
   * 
   */
  static extractBookIdFromUrl(bookUrl) {
    const urlRegExp = new RegExp("https://docs.google.com/spreadsheets/d/([^/]+)/");
    const matchResult = bookUrl.match(urlRegExp);
    let bookId;
    
    if(matchResult && matchResult.length >= 2) {
      bookId = matchResult[1];  // キャプチャグループは2つ目の要素に入っているので1を指定
    }
    return bookId;
  }

  /** 
   * 指定ブックをコピーして新しいブックを作成する
   * 
   * @param {string}              copyFromBookId コピー元BookId
   * @param {BookControllerParam} param コピー先BookControllerのパラメータ
   * 
   * @return {BookController} コピーして作られたブックのBookController
   * 
   */
  static createBookFromAnotherBook(copyFromBookId, param) {
    try {
      const copyFromBook = DriveApp.getFileById(copyFromBookId)
      const copyToBook = copyFromBook.makeCopy();
      const controller = new BookController({
        bookId: copyToBook.getId(),
        bookName: param.bookName || copyFromBook.getName(),
        folderId: param.folderId
      });

      return controller;
    } catch(e) {
      console.log(e.stack);
      throw new Error("ブックのコピーに失敗しました");
    }
  }
}
