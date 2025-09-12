/**
 * Web App Entrypoint + テンプレート読み分け
 */
function doGet(e) {
  const id   = e && e.parameter && e.parameter.id   ? e.parameter.id   : '';
  const view = e && e.parameter && e.parameter.view ? e.parameter.view : '';

  // データロード
  const data = id ? getIncidentById(id) : {};

  // 部署リストをプロパティから取得
  const props = PropertiesService.getScriptProperties();
  const deptsJson = props.getProperty('DEPT_WEBHOOKS_JSON') || '{}';
  const depts = Object.keys(JSON.parse(deptsJson));

  // OJT リンク/画像（プロパティ化）
  const ojt = {
    link: props.getProperty('OJT_LINK_URL') || '',
    img:  props.getProperty('OJT_IMAGE_URL') || ''
  };

  // 画面切替（report → Report.html）
  if (view === 'report') {
    const t = HtmlService.createTemplateFromFile('Report');
    t.data = data; t.ojt = ojt;
    return t.evaluate().setTitle(`レポート: ${id}`).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const t = HtmlService.createTemplateFromFile('index');
  t.data = data; t.ojt = ojt;
  t.departments = depts;
  return t.evaluate().setTitle('リスクアセスメント UI').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * HTMLテンプレート内で別ファイルをインクルードするための関数
 * 呼び出し名と実際のファイル名をマッピングして互換性を維持
 */
function include(filename) {
  let actualFilename = filename;
  switch (filename) {
    case 'CSS':
      actualFilename = 'Stylesheet';
      break;
    case 'javascript':
      actualFilename = 'JavaScript';
      break;
    case 'Reporthtml':
      actualFilename = 'Report';
      break;
  }
  return HtmlService.createHtmlOutputFromFile(actualFilename).getContent();
}

