// ====== Spreadsheet Settings ======
const SPREADSHEET_ID = "1wNkohKK2kZKXUA6zekJvCDivFdU3rZBg8xJovWEqfSE";
const SHEET_NAME     = "Input";
const TOKEN_SHEET_NAME = 'Token';

// ====== Script Properties (To be set in Project Settings > Script Properties) ======
const SCRIPT_PROPS       = PropertiesService.getScriptProperties();
const API_KEY            = SCRIPT_PROPS.getProperty('API_KEY');
const WEBAPP_BASE_URL    = SCRIPT_PROPS.getProperty('WEBAPP_BASE_URL') || "";
const DEPT_WEBHOOKS_JSON = SCRIPT_PROPS.getProperty('DEPT_WEBHOOKS_JSON') || "{}";

// ====== Gemini AI Model ======
const GEMINI_MODEL = 'gemini-1.5-flash-latest';

