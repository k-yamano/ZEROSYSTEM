/**
 * Webアプリケーションにアクセスがあった際のメインの処理（エントリーポイント）です。
 * URLのIDに基づいてデータを取得し、HTMLテンプレートに渡してUIを生成します。
 * @param {object} e - Webアプリへのリクエスト情報を含むイベントオブジェクト
 * @returns {HtmlOutput} - 表示するHTMLページ
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
    // ★★★ HTMLテンプレート側で 'data' という変数名が使えるように設定 ★★★
    template.data = incidentData;
    
    // 部署のリストをテンプレートに渡す
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
 * @param {string} filename - 読み込むファイル名（拡張子なし）
 * @returns {string} - ファイルのコンテンツ
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 日付オブジェクトを指定された書式の文字列に変換します。
 * @param {Date | string} date - 変換する日付
 * @param {string} format - 出力する日付の書式
 * @returns {string} - フォーマットされた日付文字列
 */
function formatDate(date, format = 'yyyy/MM/dd HH:mm') {
    if(!date) return '';
    try {
      return Utilities.formatDate(new Date(date), 'JST', format);
    } catch(e) {
      return ''; // 無効な日付の場合は空文字を返す
    }
}
