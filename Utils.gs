/**
 * Generates a unique ID (e.g., AB-C12).
 * @returns {string} The generated unique ID.
 */
function newId() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const alphaNum = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let prefix = '', suffix = '';
  for (let i = 0; i < 2; i++) prefix += chars.charAt(Math.floor(Math.random() * chars.length));
  for (let i = 0; i < 3; i++) suffix += alphaNum.charAt(Math.floor(Math.random() * alphaNum.length));
  return `${prefix}-${suffix}`;
}

/**
 * Formats a date object into a string.
 * @param {Date|string} date The date to format.
 * @param {string} [format='yyyy/MM/dd HH:mm'] The desired format.
 * @returns {string} The formatted date string.
 */
function formatDate(date, format = 'yyyy/MM/dd HH:mm') {
    if(!date) return '';
    return Utilities.formatDate(new Date(date), 'JST', format);
}

