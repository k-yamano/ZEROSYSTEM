/**
 * ★追加：OJT IDをキーにして、別のスプレッドシートからOJTの登録内容を取得する
 * @param {string} ojtId - 検索対象のOJT ID
 * @returns {string} - 見つかったOJT項目、見つからなければ空文字
 */
function getOjtDetails_(ojtId) {
  if (!ojtId) return "";
  try {
    const ojtSheetId = "1sXRUGEviaDGFwu2OAiiNPz9wMaxYsy-gdohBmDxUK7k"; // OJT登録シートのID
    const ojtSheet = SpreadsheetApp.openById(ojtSheetId).getSheets()[0]; // 最初のシートを対象とする
    const data = ojtSheet.getDataRange().getValues();
    const headers = data.shift(); // ヘッダー行を取得

    // M列（ID）とD列（OJT項目）のインデックスを取得
    const idColIndex = headers.indexOf("ID");
    const detailsColIndex = headers.indexOf("OJT項目を記入してください");

    if (idColIndex === -1 || detailsColIndex === -1) {
      Logger.log("OJTシートのヘッダーが見つかりません（ID または OJT項目を記入してください）");
      return "（OJTシートの列が見つかりません）";
    }

    // OJT IDに一致する行を探す
    for (let i = 0; i < data.length; i++) {
      if (data[i][idColIndex] == ojtId) {
        return data[i][detailsColIndex]; // D列の内容を返す
      }
    }
    return "（該当するOJT情報なし）";
  } catch (e) {
    Logger.log(`OJT詳細の取得中にエラーが発生しました: ${e.message}`);
    return `（エラー: ${e.message}）`;
  }
}

/**
 * Webアプリケーションにアクセスがあった際のメインの処理（エントリーポイント）です。
 */
function doGet(e) {
  try {
    const id = (e.parameter.id || '').toString().trim();
    if (!id) {
      return HtmlService.createHtmlOutput('<h1>エラー: IDが指定されていません。</h1>');
    }

    const incidentData = getDataById(id);
    if (!incidentData || Object.keys(incidentData).length === 0) {
      return HtmlService.createHtmlOutput(`<h1>エラー: ID「${id}」のデータが見つかりません。</h1>`);
    }

    // ★修正: OJT詳細を取得してデータに追加
    if (incidentData.ojt_id) {
      incidentData.ojt_details = getOjtDetails_(incidentData.ojt_id);
    } else {
      incidentData.ojt_details = "（OJT ID未登録）";
    }

    const view = (e.parameter.view || '').toString();

    // レポート（印刷用）画面
    if (view === 'report') {
      const template = HtmlService.createTemplateFromFile('Report');
      template.data = incidentData;
      return template.evaluate()
        .setTitle(`レポート: ${id}`)
        .addMetaTag('viewport','width=device-width, initial-scale=1');
    }

    // 通常の操作画面
    const template = HtmlService.createTemplateFromFile('index');
    template.data = incidentData;
    
    const props = PropertiesService.getScriptProperties();
    const deptsJson = props.getProperty('DEPT_WEBHOOKS_JSON') || '{}';
    template.departments = Object.keys(JSON.parse(deptsJson));

    return template.evaluate()
      .setTitle(`リスク評価: ${id}`)
      .addMetaTag('viewport','width=device-width, initial-scale=1');

  } catch (err) {
    Logger.log(`doGet Error: ${err.message}\n${err.stack}`);
    return HtmlService.createHtmlOutput(`<h1>サーバーエラーが発生しました</h1><p>${err.message}</p>`);
  }
}

/**
 * HTMLテンプレート内で別のHTMLファイル（CSSやJavaScript）を読み込むためのヘルパー関数です。
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * HTMLテンプレートから呼び出すための日付フォーマット関数
 */
function formatDateForHtml(date, format = 'yyyy/MM/dd HH:mm') {
    if(!date) return '';
    try {
      return Utilities.formatDate(new Date(date), 'JST', format);
    } catch(e) {
      return '';
    }
}

