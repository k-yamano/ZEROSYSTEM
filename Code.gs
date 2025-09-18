/** Web Appのエントリーポイント */
function doGet(e) {
  try {
    const id = (e.parameter.id || '').toString().trim();
    if (!id) return HtmlService.createHtmlOutput('<h1>エラー: IDが指定されていません。</h1>');

    // ★★★ エラーの修正：呼び出す関数名を getDataById に変更 ★★★
    const incident = getDataById(id); 
    if (!incident) return HtmlService.createHtmlOutput(`<h1>エラー: ID「${id}」が見つかりません。</h1>`);

    const view = (e.parameter.view || '').toString();
    // レポート画面
    if (view === 'report') {
      const t = HtmlService.createTemplateFromFile('Report');
      t.data = incident;
      return t.evaluate()
        .setTitle(`レポート: ${id}`)
        .addMetaTag('viewport','width=device-width, initial-scale=1');
    }

    // 通常のインタラクティブ画面
    const t = HtmlService.createTemplateFromFile('index');
    t.data = incident;
    // ★旧版と同様に部署リストを渡す
    const props = PropertiesService.getScriptProperties();
    const deptsJson = props.getProperty('DEPT_WEBHOOKS_JSON') || '{}';
    t.departments = Object.keys(JSON.parse(deptsJson));
    return t.evaluate()
      .setTitle(`リスク評価: ${id}`)
      .addMetaTag('viewport','width=device-width, initial-scale=1');
  } catch (err) {
    Logger.log(`doGet Error: ${err.toString()}\n${err.stack}`);
    return HtmlService.createHtmlOutput(`<h1>エラー</h1><p>${err.message}</p>`);
  }
}

/** HTML内で別ファイルをインクルードする */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}