/** Web Appのエントリーポイント */
function doGet(e) {
  const id = e?.parameter?.id || '';
  const view = e?.parameter?.view || '';
  const data = id ? getDataById(id) : {};
  const props = PropertiesService.getScriptProperties();
  const deptsJson = props.getProperty('DEPT_WEBHOOKS_JSON') || '{}';
  const depts = Object.keys(JSON.parse(deptsJson));

  if (view === 'report' && id) {
    const t = HtmlService.createTemplateFromFile('Report');
    t.data = data;
    return t.evaluate().setTitle(`レポート: ${id}`).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const t = HtmlService.createTemplateFromFile('index');
  t.data = data;
  t.departments = depts;
  return t.evaluate().setTitle('リスクアセスメント UI').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** HTML内で別ファイルをインクルードする */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}