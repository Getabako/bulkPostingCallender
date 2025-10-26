// フォルダIDを定数として定義
const FOLDER_ID = '1b58Ox5XhShFc7EJl-vp2WWmWFde_D9Fn';
// サンクスメッセージ画像のIDを定数として定義
const THANKS_IMAGE_ID = '10I9rmtmLEDedqKA7QPK9G7_aNKTlADgV';

/**
 * スプレッドシート起動時にカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('画像リンク')
    .addItem('画像リンクを挿入', 'insertImageLinks')
    .addToUi();
}

/**
 * C列が空の場合、Google Driveフォルダ内の画像リンクを挿入
 * 4つの画像+サンクスメッセージ画像の5枚セットをカンマ区切りでC列に配置（ファイル名順）
 */
function insertImageLinks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  // A列とC列の値をチェック（2行目以降）
  if (lastRow > 1) {
    const aColumnValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const cColumnValues = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
    const hasAContent = aColumnValues.some(row => row[0] !== '');
    const hasCContent = cColumnValues.some(row => row[0] !== '');

    if (hasAContent || hasCContent) {
      SpreadsheetApp.getUi().alert('A列またはC列にすでにデータが入力されています。\n処理を中止します。');
      return;
    }
  }

  try {
    // フォルダから画像ファイルを取得
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const files = folder.getFiles();
    const imageFiles = [];

    // 画像ファイルを収集
    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();

      // 画像ファイルのみフィルタ
      if (mimeType.startsWith('image/')) {
        imageFiles.push({
          name: file.getName(),
          id: file.getId()
        });
      }
    }

    // 画像がない場合
    if (imageFiles.length === 0) {
      SpreadsheetApp.getUi().alert('フォルダ内に画像ファイルが見つかりませんでした。');
      return;
    }

    // ファイル名でソート（数字を含むファイル名を正しくソート）
    imageFiles.sort((a, b) => {
      return a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
    });

    // 画像リンクを生成（Googleキャッシュサーバー経由の直接リンク形式）
    const imageLinks = imageFiles.map(file => {
      return `https://lh3.googleusercontent.com/d/${file.id}`;
    });

    // サンクスメッセージ画像のリンクを生成
    const thanksImageLink = `https://lh3.googleusercontent.com/d/${THANKS_IMAGE_ID}`;

    // 4つずつのグループに分割し、各グループにサンクスメッセージ画像を追加してカンマ区切りで連結
    const rowData = [];
    for (let i = 0; i < imageLinks.length; i += 4) {
      const imageSet = [
        imageLinks[i] || '',
        imageLinks[i + 1] || '',
        imageLinks[i + 2] || '',
        imageLinks[i + 3] || '',
        thanksImageLink
      ];
      // 空でない画像リンクのみをフィルタしてカンマ区切りで連結
      const combinedLinks = imageSet.filter(link => link !== '').join(',');
      rowData.push([combinedLinks]);
    }

    // 日付データを作成（今日から開始し、毎日18:00に設定）
    const dateData = [];
    const startDate = new Date();
    startDate.setHours(18, 0, 0, 0); // 時刻を18:00:00に設定

    for (let i = 0; i < rowData.length; i++) {
      const currentDate = new Date(startDate);
      currentDate.setDate(startDate.getDate() + i); // i日後の日付

      // 日付を「MM/dd/yyyy HH:mm」形式の文字列にフォーマット
      const month = String(currentDate.getMonth() + 1).padStart(2, '0');
      const day = String(currentDate.getDate()).padStart(2, '0');
      const year = currentDate.getFullYear();
      const hours = String(currentDate.getHours()).padStart(2, '0');
      const minutes = String(currentDate.getMinutes()).padStart(2, '0');
      const formattedDate = `${month}/${day}/${year} ${hours}:${minutes}`;

      dateData.push([formattedDate]);
    }

    // 1行目にヘッダーを追加
    sheet.getRange(1, 1, 1, 3).setValues([['Date', 'Text', 'Media URL(s)']]);

    // A列に日付、C列に画像リンクを挿入（2行目から）
    sheet.getRange(2, 1, dateData.length, 1).setValues(dateData);
    sheet.getRange(2, 3, rowData.length, 1).setValues(rowData);

    const groupCount = rowData.length;
    const totalImages = imageLinks.length;
    SpreadsheetApp.getUi().alert(
      `${totalImages}件の画像リンクを挿入しました。\n（${groupCount}グループ、各グループ最大4画像+サンクスメッセージ画像）\n\n順序: ファイル名順`
    );

  } catch (error) {
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${error.message}\n\nフォルダへのアクセス権限を確認してください。`);
  }
}
