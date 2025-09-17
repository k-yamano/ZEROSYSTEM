/** ユニークIDを生成する (例: AB-C12) */
function newId() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const alphaNum = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz012356789';
  let prefix = '', suffix = '';
  for (let i = 0; i < 2; i++) prefix += chars.charAt(Math.floor(Math.random() * chars.length));
  for (let i = 0; i < 3; i++) suffix += alphaNum.charAt(Math.floor(Math.random() * alphaNum.length));
  return `${prefix}-${suffix}`;
}

/** 日付を指定のフォーマットで文字列化する */
function formatDate(date, format = 'yyyy/MM/dd HH:mm') {
    if(!date) return '';
    return Utilities.formatDate(new Date(date), 'JST', format);
}