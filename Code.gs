/**
 * Web App Entrypoint + Template routing
 */
function doGet(e) {
  const id   = e && e.parameter && e.parameter.id   ? e.parameter.id   : '';
  const view = e && e.parameter && e.parameter.view ? e.parameter.view : '';

  const data = id ? getIncidentById(id) : {};

  const props = PropertiesService.getScriptProperties();
  const deptsJson = props.getProperty('DEPT_WEBHOOKS_JSON') || '{}';
  const depts = Object.keys(JSON.parse(deptsJson));

  if (view === 'report') {
    const t = HtmlService.createTemplateFromFile('Report');
    t.data = data;
    return t.evaluate().setTitle(`レポート: ${id}`).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const t = HtmlService.createTemplateFromFile('index');
  t.data = data;
  t.departments = depts;
  return t.evaluate().setTitle('リスクアセスメント UI').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Function to include other HTML files within a template,
 * mapping old names to new ones for compatibility.
 */
function include(filename) {
  let actualFilename = filename;
  if (filename === 'CSS') {
    actualFilename = 'Stylesheet';
  } else if (filename === 'javascript') {
    actualFilename = 'JavaScript';
  }
  return HtmlService.createHtmlOutputFromFile(actualFilename).getContent();
}

