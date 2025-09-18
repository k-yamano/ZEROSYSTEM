// ====== Spreadsheet Settings ======
const SPREADSHEET_ID = "1wNkohKK2kZKXUA6zekJvCDivFdU3rZBg8xJovWEqfSE";
const SHEET_NAME     = "Input";
const TOKEN_SHEET    = "Token";
const LOG_SHEET      = "Log";
// ====== Script Properties ======
const SCRIPT_PROPS       = PropertiesService.getScriptProperties();
const API_KEY            = SCRIPT_PROPS.getProperty('API_KEY');
const WEBAPP_BASE_URL    = SCRIPT_PROPS.getProperty('WEBAPP_BASE_URL');

// ====== Gemini AI Model ======
const GEMINI_MODEL = 'gemini-1.5-flash-latest';
// ====== Header to Key Mapping ======
const KEY_MAP = {
  'ID': 'unique_id', 'タイムスタンプ': 'timestamp', 'メールアドレス': 'discoverer_email', '件名': 'subject',
  'ヒヤリハット内容': 'details', '希望する評価者': 'hopeful_evaluator', 'ステータス': 'status', '担当部署': 'department',
  '期限': 'deadline', '暫定予算': 'provisional_budget', '評価者によるコメント': 'secondary_eval_comment', // AC列
  '評価理由': 'evaluation_reason', '改善プラン': 'improvement_plan_ai', '頻度（AI評価）': 'frequency_ai',
  '発生の可能性（AI評価）': 'likelihood_ai', '重篤度（AI評価）': 'severity_ai', 'リスクの見積もり（AI評価）': 'risk_score_ai',
  '優先順位': 'priority', '事故の型分類（AI評価）': 'accident_type_ai', '起因物（AI評価）': 'causal_agent_ai',
  '評価者によるリスク修正': 'modified_by_evaluator_primary', '改善完了報告': 'improvement_details',
  '報告者': 'reporter', '協力者1': 'team_member_2', '協力者2': 'team_member_3', 'URL': 'reference_url',
  'OJT': 'ojt_completed', 'OJTID': 'ojt_id', '費用': 'cost', '工数': 'effort', '効果': 'effect',
  'AI改善評価': 'improvement_ai_comment', '頻度（改善評価）': 'post_frequency',
  '発生の可能性（改善後評価）': 'post_likelihood', '重篤度（改善後評価）': 'post_severity',
  '改善後のリスクの見積もり': 'post_risk', // AU列に相当
  '改善後の優先度': 'post_priority', // AW列に相当
  '評価者による修正（最終）': 'modified_by_evaluator_final',
  'リスク低減値': 'risk_reduction_value',   // AY列
  '最終評価コメント': 'final_eval_comment',
  '最終通知日時': 'final_notification_timestamp', // AP列
  '差し戻し理由': 'revert_reason'
};

