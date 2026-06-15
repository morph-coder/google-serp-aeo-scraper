/**
 * Google SERP & AEO Scraper — Google Sheets template
 *
 * Setup:
 * 1. Extensions → Apps Script → paste this file + appsscript.json
 * 2. Reload spreadsheet → first open auto-creates sheets & headers
 * 3. SERP Tools → Configure Apify token
 * 4. Add keywords on Keywords sheet, adjust Settings
 * 5. SERP Tools → Run SERP scan (results fetch via 1-min trigger)
 */

const SHEET_SETTINGS = 'Settings';
const SHEET_KEYWORDS = 'Keywords';
const SHEET_RESULTS = 'Results';
const SHEET_LLM = 'LLM Summary';
const SHEET_LLM_ANSWERS = 'LLM Answers';
const SHEET_LLM_CITATIONS = 'LLM Citations';
const SHEET_RUN_LOG = 'Run Log';

const PROP_APIFY_TOKEN = 'APIFY_TOKEN';
const PROP_LAST_RUN_ID = 'LAST_RUN_ID';
const PROP_POLL_ATTEMPT = 'POLL_ATTEMPT';
const PROP_TEMPLATE_INITIALIZED = 'TEMPLATE_INITIALIZED';

const TRIGGER_HANDLER = 'fetchScheduledRunResults_';
const POLL_INTERVAL_MS = 60 * 1000;
const MAX_POLL_ATTEMPTS = 30;

const DEFAULT_ACTOR_ID = 'morph_coder/google-serp-aeo-scraper';

const FLAT_COLUMNS = [
  'keyword',
  'resultsTotal',
  'rank',
  'serpSlot',
  'queryPage',
  'position',
  'title',
  'url',
  'displayedUrl',
  'description',
  'isPaid',
  'type',
  'targetCountry',
  'device',
  'serpUrl',
  'fetchedAt',
  'llmChatGptCited',
  'llmChatGptUrlRank',
  'llmChatGptDomainRank',
  'llmGeminiCited',
  'llmGeminiUrlRank',
  'llmGeminiDomainRank',
  'llmPerplexityCited',
  'llmPerplexityUrlRank',
  'llmPerplexityDomainRank',
];

const LLM_PROVIDER_EXPORTS = [
  { resultKey: 'chatGptSearchResult', prefix: 'chatGpt', label: 'ChatGPT' },
  { resultKey: 'geminiSearchResult', prefix: 'gemini', label: 'Gemini' },
  { resultKey: 'perplexitySearchResult', prefix: 'perplexity', label: 'Perplexity' },
  { resultKey: 'copilotSearchResult', prefix: 'copilot', label: 'Copilot' },
  { resultKey: 'deepSeekSearchResult', prefix: 'deepSeek', label: 'DeepSeek' },
];

const LLM_SUMMARY_META_COLUMNS = [
  'keyword',
  'targetCountry',
  'queryPage',
  'fetchedAt',
  'overlapPercent',
  'primaryDomain',
  'googleRank',
  'llmRank',
  'rankDelta',
];

/** Max chars per cell (Google Sheets limit is 50k). */
const LLM_CELL_CHAR_LIMIT = 45000;

function getLlmSummaryColumns_() {
  const columns = LLM_SUMMARY_META_COLUMNS.slice();
  LLM_PROVIDER_EXPORTS.forEach(function (provider) {
    columns.push(provider.prefix + '_query');
    columns.push(provider.prefix + '_webQuery');
    columns.push(provider.prefix + '_citationCount');
    columns.push(provider.prefix + '_error');
  });
  return columns;
}

const LLM_ANSWERS_COLUMNS = [
  'keyword',
  'targetCountry',
  'queryPage',
  'fetchedAt',
  'provider',
  'query',
  'webSearchQuery',
  'model',
  'answer',
  'citationCount',
  'error',
];

const LLM_CITATIONS_COLUMNS = [
  'keyword',
  'targetCountry',
  'queryPage',
  'fetchedAt',
  'provider',
  'query',
  'webSearchQuery',
  'rank',
  'url',
  'title',
];

const RUN_LOG_COLUMNS = [
  'loggedAt',
  'runId',
  'status',
  'organicUrlsCharged',
  'chatGptCalls',
  'geminiCalls',
  'perplexityCalls',
  'deepSeekCalls',
  'pagesFetched',
  'costUsd',
  'note',
];

/** ISO country codes supported by the Actor (matches input_schema). */
const COUNTRY_CODES = [
  'us', 'gb', 'de', 'fr', 'es', 'it', 'in', 'au', 'ca', 'br', 'mx', 'jp', 'kr',
  'nl', 'pl', 'ua', 'ae', 'sg', 'hk', 'tw', 'se', 'no', 'dk', 'fi', 'be', 'at',
  'ch', 'cz', 'ro', 'pt', 'gr', 'tr', 'il', 'sa', 'za', 'ng', 'ph', 'id', 'my',
  'th', 'vn', 'nz', 'ie', 'ar', 'cl', 'co', 'pe',
];

const MAX_ORGANIC_OPTIONS = ['10', '20', '30', '50', '100'];
const MAX_PAGES_OPTIONS = ['1', '2', '3', '5', '10'];
const MAX_TOTAL_OPTIONS = ['(none)', '20', '50', '100', '200', '500', '1000'];
const OUTPUT_FORMAT_OPTIONS = ['page', 'flat', 'both'];
const QUICK_DATE_OPTIONS = ['(none)', 'd', 'd10', 'w2', 'm6', 'y1'];
const LLM_SCOPE_OPTIONS = ['perQuery', 'perPage'];
const LLM_QUERY_MODE_OPTIONS = ['sameAsKeyword', 'conversational'];
const LLM_APPLY_TO_OPTIONS = ['brandLike', 'targetOnly', 'all'];
const PERPLEXITY_RECENCY_OPTIONS = ['day', 'week', 'month', 'year'];
const BOOL_OPTIONS = ['TRUE', 'FALSE'];

/** Row index in Settings column B for each field key (rebuilt on setup / read). */
var SETTING_ROW = {};

// ─── Menu & first-run setup ───────────────────────────────────────────────

function onOpen() {
  initializeTemplateIfNeeded_();
  buildMenu_();
}

function onInstall() {
  initializeTemplate_();
  buildMenu_();
}

function buildMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('SERP Tools')
    .addItem('Configure Apify token', 'configureApifyToken')
    .addItem('Run SERP scan', 'runSerpScan')
    .addItem('Fetch last run results', 'fetchLastRunResults')
    .addSeparator()
    .addItem('Initialize / reset template', 'resetTemplate')
    .addItem('Clear results sheets', 'clearResults')
    .addToUi();
}

function initializeTemplateIfNeeded_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_TEMPLATE_INITIALIZED) === 'true') {
    return;
  }
  initializeTemplate_();
}

function resetTemplate() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Reset template?',
    'This recreates sheet headers and default settings. Keywords on the Keywords sheet are kept.',
    ui.ButtonSet.YES_NO,
  );
  if (answer !== ui.Button.YES) {
    return;
  }
  PropertiesService.getScriptProperties().setProperty(PROP_TEMPLATE_INITIALIZED, 'false');
  initializeTemplate_();
  ui.alert('Template reset. Review Settings and Keywords, then run a scan.');
}

function initializeTemplate_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  setupSettingsSheet_(ss);
  setupKeywordsSheet_(ss);
  setupResultsSheet_(ss);
  setupLlmSummarySheet_(ss);
  setupLlmAnswersSheet_(ss);
  setupLlmCitationsSheet_(ss);
  setupRunLogSheet_(ss);

  reorderSheets_(ss);
  removeEmptyDefaultSheet_(ss);

  PropertiesService.getScriptProperties().setProperty(PROP_TEMPLATE_INITIALIZED, 'true');
  setStatusNote_('Template ready. Configure Apify token, then Run SERP scan.');
}

function getSettingsFields_() {
  return [
    { type: 'section', label: 'SERP' },
    { key: 'countryCode', label: 'Country code (ISO)', default: 'in', validation: 'country' },
    { key: 'maxResultsPerQuery', label: 'Max organic per keyword', default: '10', validation: 'maxOrganic' },
    { key: 'maxPagesPerQuery', label: 'Max pages per query', default: '2', validation: 'maxPages' },
    { key: 'maxTotalResults', label: 'Max total organic (whole run)', default: '(none)', validation: 'maxTotal' },
    { key: 'outputFormat', label: 'Output format', default: 'flat', validation: 'outputFormat' },
    { key: 'mobileResults', label: 'Mobile results', default: 'FALSE', validation: 'bool' },
    { key: 'disableGoogleSearchResults', label: 'Disable Google SERP (LLM only)', default: 'FALSE', validation: 'bool' },
    { type: 'section', label: 'Search filters' },
    { key: 'forceExactMatch', label: 'Force exact match', default: 'FALSE', validation: 'bool' },
    { key: 'site', label: 'Site filter (site:)', default: '' },
    { key: 'relatedToSite', label: 'Related to site', default: '' },
    { key: 'wordsInTitle', label: 'Words in title (comma-separated)', default: '' },
    { key: 'wordsInText', label: 'Words in text (comma-separated)', default: '' },
    { key: 'wordsInUrl', label: 'Words in URL (comma-separated)', default: '' },
    { key: 'fileTypes', label: 'File types (comma-separated)', default: '' },
    { key: 'quickDateRange', label: 'Quick date range', default: '(none)', validation: 'quickDate' },
    { key: 'afterDate', label: 'After date (YYYY-MM-DD)', default: '' },
    { key: 'beforeDate', label: 'Before date (YYYY-MM-DD)', default: '' },
    { key: 'searchLanguage', label: 'Search language (lr)', default: '' },
    { key: 'languageCode', label: 'UI language (hl)', default: '' },
    { key: 'locationUule', label: 'Location UULE', default: '' },
    { type: 'section', label: 'Ads & debug' },
    { key: 'focusOnPaidAds', label: 'Focus on paid ads', default: 'FALSE', validation: 'bool' },
    { key: 'includeUnfilteredResults', label: 'Include unfiltered results', default: 'FALSE', validation: 'bool' },
    { key: 'saveHtml', label: 'Save HTML to dataset', default: 'FALSE', validation: 'bool' },
    { key: 'saveHtmlToKeyValueStore', label: 'Save HTML to key-value store', default: 'TRUE', validation: 'bool' },
    { key: 'includeIcons', label: 'Include favicon icons', default: 'FALSE', validation: 'bool' },
    { type: 'section', label: 'LLM add-ons' },
    { key: 'llmSearchScope', label: 'LLM search scope', default: 'perQuery', validation: 'llmScope' },
    { key: 'llmQueryMode', label: 'LLM query mode', default: 'sameAsKeyword', validation: 'llmQueryMode' },
    { key: 'llmApplyTo', label: 'Run LLM for', default: 'all', validation: 'llmApplyTo' },
    { key: 'targetBrand', label: 'Target brand', default: '' },
    { key: 'targetDomains', label: 'Target domains (comma-separated)', default: '' },
    { key: 'enableChatGpt', label: 'Enable ChatGPT', default: 'TRUE', validation: 'bool' },
    { key: 'enableGemini', label: 'Enable Gemini', default: 'TRUE', validation: 'bool' },
    { key: 'enablePerplexity', label: 'Enable Perplexity', default: 'TRUE', validation: 'bool' },
    { key: 'perplexitySearchRecency', label: 'Perplexity search recency', default: 'week', validation: 'perplexityRecency' },
    { key: 'perplexityReturnImages', label: 'Perplexity return images', default: 'FALSE', validation: 'bool' },
    { key: 'perplexityReturnRelatedQuestions', label: 'Perplexity return related questions', default: 'FALSE', validation: 'bool' },
    { key: 'enableCopilot', label: 'Enable Copilot', default: 'FALSE', validation: 'bool' },
    { key: 'enableDeepSeek', label: 'Enable DeepSeek (experimental)', default: 'TRUE', validation: 'bool' },
    { key: 'enableAiMode', label: 'Enable AI Mode (coming soon)', default: 'TRUE', validation: 'bool' },
  ];
}

function rebuildSettingRows_() {
  SETTING_ROW = {};
  let row = 2;
  getSettingsFields_().forEach(function (field) {
    if (field.type === 'section') {
      row++;
      return;
    }
    if (field.key) {
      SETTING_ROW[field.key] = row;
    }
    row++;
  });
  SETTING_ROW.status = row;
}

function setupSettingsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEET_SETTINGS);
  unmergeSettingsSheet_(sheet);
  sheet.clear();
  sheet.getRange('A1:B1').setValues([['Setting', 'Value']]).setFontWeight('bold');

  const fields = getSettingsFields_();
  let row = 2;
  fields.forEach(function (field) {
    if (field.type === 'section') {
      sheet
        .getRange(row, 1, row, 2)
        .setBackground('#f3f3f3')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      sheet.getRange(row, 1).setValue('— ' + field.label + ' —');
      row++;
      return;
    }
    sheet.getRange(row, 1).setValue(field.label).setFontWeight('bold');
    sheet.getRange(row, 2).setValue(field.default);
    if (field.validation) {
      applySettingValidation_(sheet, row, field.validation);
    }
    row++;
  });

  sheet.getRange(row, 1).setValue('Status').setFontWeight('bold');
  sheet.getRange(row, 2).setValue('');
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 200);
  sheet.setFrozenRows(1);
  rebuildSettingRows_();
}

/** breakApart before clear — merged cells survive sheet.clear() and break later writes. */
function unmergeSettingsSheet_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 60);
  try {
    sheet.getRange(1, 1, lastRow, 2).breakApart();
  } catch (e) {
    // no merges to break
  }
}

function applySettingValidation_(sheet, row, type) {
  const cell = sheet.getRange(row, 2);
  switch (type) {
    case 'country':
      cell.setDataValidation(listValidation_(COUNTRY_CODES));
      break;
    case 'maxOrganic':
      cell.setDataValidation(listValidation_(MAX_ORGANIC_OPTIONS));
      break;
    case 'maxPages':
      cell.setDataValidation(listValidation_(MAX_PAGES_OPTIONS));
      break;
    case 'maxTotal':
      cell.setDataValidation(listValidation_(MAX_TOTAL_OPTIONS));
      break;
    case 'outputFormat':
      cell.setDataValidation(listValidation_(OUTPUT_FORMAT_OPTIONS));
      break;
    case 'quickDate':
      cell.setDataValidation(listValidation_(QUICK_DATE_OPTIONS));
      break;
    case 'llmScope':
      cell.setDataValidation(listValidation_(LLM_SCOPE_OPTIONS));
      break;
    case 'llmQueryMode':
      cell.setDataValidation(listValidation_(LLM_QUERY_MODE_OPTIONS));
      break;
    case 'llmApplyTo':
      cell.setDataValidation(listValidation_(LLM_APPLY_TO_OPTIONS));
      break;
    case 'perplexityRecency':
      cell.setDataValidation(listValidation_(PERPLEXITY_RECENCY_OPTIONS));
      break;
    case 'bool':
      cell.setDataValidation(listValidation_(BOOL_OPTIONS));
      break;
    default:
      break;
  }
}

function listValidation_(values) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
}

function setupKeywordsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEET_KEYWORDS);
  const lastRow = sheet.getLastRow();
  const existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : null;

  sheet.clear();
  sheet.getRange('A1').setValue('Keyword').setFontWeight('bold');
  sheet.setColumnWidth(1, 280);

  if (existing && existing.flat().some(function (v) { return String(v).trim(); })) {
    sheet.getRange(2, 1, existing.length, 1).setValues(existing);
  } else {
    sheet.getRange('A2:A4').setValues([['nike'], ['tesla'], ['NASA']]);
  }
}

function setupResultsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEET_RESULTS);
  sheet.clear();
  writeHeaderRow_(sheet, FLAT_COLUMNS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, FLAT_COLUMNS.length).setBackground('#e8f0fe');
}

function setupLlmSummarySheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEET_LLM);
  const columns = getLlmSummaryColumns_();
  sheet.clear();
  writeHeaderRow_(sheet, columns);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, columns.length).setBackground('#e6f4ea');
}

function setupLlmAnswersSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEET_LLM_ANSWERS);
  sheet.clear();
  writeHeaderRow_(sheet, LLM_ANSWERS_COLUMNS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, LLM_ANSWERS_COLUMNS.length).setBackground('#d9ead3');
}

function setupLlmCitationsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEET_LLM_CITATIONS);
  sheet.clear();
  writeHeaderRow_(sheet, LLM_CITATIONS_COLUMNS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, LLM_CITATIONS_COLUMNS.length).setBackground('#fff2cc');
}

function setupRunLogSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEET_RUN_LOG);
  if (sheet.getLastRow() === 0) {
    writeHeaderRow_(sheet, RUN_LOG_COLUMNS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, RUN_LOG_COLUMNS.length).setBackground('#fce8e6');
  }
}

function reorderSheets_(ss) {
  const order = [
    SHEET_SETTINGS,
    SHEET_KEYWORDS,
    SHEET_RESULTS,
    SHEET_LLM,
    SHEET_LLM_ANSWERS,
    SHEET_LLM_CITATIONS,
    SHEET_RUN_LOG,
  ];
  order.forEach(function (name, index) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    }
  });
  const settings = ss.getSheetByName(SHEET_SETTINGS);
  if (settings) {
    ss.setActiveSheet(settings);
  }
}

function removeEmptyDefaultSheet_(ss) {
  const sheets = ss.getSheets();
  if (sheets.length <= 7) {
    return;
  }
  const defaultSheet = sheets.find(function (s) {
    return s.getName() === 'Sheet1' || s.getName() === 'Аркуш1';
  });
  if (defaultSheet && defaultSheet.getLastRow() <= 1 && defaultSheet.getLastColumn() <= 1) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch (e) {
      // ignore if last sheet
    }
  }
}

function writeHeaderRow_(sheet, columns) {
  sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
  sheet.getRange(1, 1, 1, columns.length).setFontWeight('bold');
}

// ─── Apify configuration ────────────────────────────────────────────────────

function configureApifyToken() {
  const ui = SpreadsheetApp.getUi();
  const tokenResp = ui.prompt(
    'Apify API token',
    'Paste your token from console.apify.com/account/integrations',
    ui.ButtonSet.OK_CANCEL,
  );
  if (tokenResp.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const token = tokenResp.getResponseText().trim();
  if (!token) {
    ui.alert('Token cannot be empty.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty(PROP_APIFY_TOKEN, token);

  ui.alert('Saved. Token is stored in Script Properties (not in cells).');
}

// ─── Run actor + scheduled fetch ────────────────────────────────────────────

function runSerpScan() {
  const ui = SpreadsheetApp.getUi();
  initializeTemplateIfNeeded_();

  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(PROP_APIFY_TOKEN);
  const actorId = DEFAULT_ACTOR_ID;

  if (!token) {
    ui.alert('Configure Apify token first (SERP Tools → Configure Apify token).');
    return;
  }

  const input = buildActorInput_();
  const queryCount = Array.isArray(input.queries) ? input.queries.length : 0;
  if (queryCount === 0) {
    ui.alert(
      'No keywords found.\n\n' +
        'Add at least one keyword on the Keywords sheet (column A, from row 2).',
    );
    return;
  }

  setStatusNote_('Starting Actor run…');
  const run = startActorRun_(token, actorId, input);
  props.setProperty(PROP_LAST_RUN_ID, run.id);
  props.deleteProperty(PROP_POLL_ATTEMPT);

  scheduleResultsFetch_();

  ui.alert(
    'Actor run started.\n\nRun ID: ' +
      run.id +
      '\n\nResults will be fetched automatically in ~1 minute (trigger). ' +
      'You can also use SERP Tools → Fetch last run results.',
  );
}

function fetchLastRunResults() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(PROP_APIFY_TOKEN);
  const runId = props.getProperty(PROP_LAST_RUN_ID);

  if (!token || !runId) {
    SpreadsheetApp.getUi().alert('No previous run found. Start a scan first.');
    return;
  }

  pollAndWriteResults_(token, runId, true);
}

/**
 * Time-based trigger entry point (do not rename without updating TRIGGER_HANDLER).
 */
function fetchScheduledRunResults_() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(PROP_APIFY_TOKEN);
  const runId = props.getProperty(PROP_LAST_RUN_ID);

  if (!token || !runId) {
    deleteTriggersForHandler_(TRIGGER_HANDLER);
    return;
  }

  pollAndWriteResults_(token, runId, false);
}

function pollAndWriteResults_(token, runId, showUi) {
  const props = PropertiesService.getScriptProperties();
  const ui = showUi ? SpreadsheetApp.getUi() : null;

  let run;
  try {
    run = apifyGet_(token, '/v2/actor-runs/' + runId);
  } catch (e) {
    appendRunLog_(runId, 'ERROR', null, 'API error: ' + e.message);
    setStatusNote_('Fetch failed — see Run Log');
    if (ui) {
      ui.alert('Failed to read run status: ' + e.message);
    }
    deleteTriggersForHandler_(TRIGGER_HANDLER);
    return;
  }

  if (run.status === 'RUNNING' || run.status === 'READY') {
    const attempt = Number(props.getProperty(PROP_POLL_ATTEMPT) || 0) + 1;
    props.setProperty(PROP_POLL_ATTEMPT, String(attempt));

    if (attempt >= MAX_POLL_ATTEMPTS) {
      setStatusNote_('Run still in progress after ' + MAX_POLL_ATTEMPTS + ' min — fetch manually');
      appendRunLog_(runId, run.status, null, 'Polling stopped — use Fetch last run results', run);
      deleteTriggersForHandler_(TRIGGER_HANDLER);
      if (ui) {
        ui.alert('Run still in progress. Try Fetch last run results later.');
      }
      return;
    }

    setStatusNote_('Run in progress… poll ' + attempt + '/' + MAX_POLL_ATTEMPTS);
    scheduleResultsFetch_();
    if (ui) {
      ui.alert('Run still in progress (' + run.status + '). Will retry in ~1 minute.');
    }
    return;
  }

  props.deleteProperty(PROP_POLL_ATTEMPT);
  deleteTriggersForHandler_(TRIGGER_HANDLER);

  if (run.status !== 'SUCCEEDED') {
    setStatusNote_('Run ' + run.status);
    appendRunLog_(runId, run.status, null, run.statusMessage || '', run);
    if (ui) {
      ui.alert('Run finished with status: ' + run.status);
    }
    return;
  }

  try {
    const counts = writeAllResults_(token, run.defaultDatasetId);
    const llmNote =
      counts.llmSource === 'flat-fallback'
        ? 'Imported ' + counts.flat + ' flat rows; LLM summary built from flat columns'
        : 'Imported ' + counts.flat + ' flat rows, ' + counts.llm + ' LLM rows';
    appendRunLog_(runId, 'SUCCEEDED', counts, llmNote, run);
    setStatusNote_(
      'Done — ' +
        counts.flat +
        ' results, ' +
        counts.llm +
        ' LLM keywords, ' +
        counts.llmCitations +
        ' citations',
    );
    if (ui) {
      ui.alert(
        'Done!\n' +
          counts.flat +
          ' flat rows\n' +
          counts.llmAnswers +
          ' LLM answer rows\n' +
          counts.llmCitations +
          ' citation rows\n\nSee Results, LLM Summary, LLM Answers, LLM Citations, Run Log.',
      );
    }
  } catch (e) {
    setStatusNote_('Import failed — see Run Log');
    appendRunLog_(runId, 'IMPORT_ERROR', null, e.message, run);
    if (ui) {
      ui.alert('Import failed: ' + e.message);
    }
  }
}

function scheduleResultsFetch_() {
  deleteTriggersForHandler_(TRIGGER_HANDLER);
  ScriptApp.newTrigger(TRIGGER_HANDLER).timeBased().after(POLL_INTERVAL_MS).create();
}

function deleteTriggersForHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function clearResults() {
  initializeTemplateIfNeeded_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupResultsSheet_(ss);
  setupLlmSummarySheet_(ss);
  setupLlmAnswersSheet_(ss);
  setupLlmCitationsSheet_(ss);
  setStatusNote_('Results cleared');
}

// ─── Actor input ──────────────────────────────────────────────────────────

function buildActorInput_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getOrCreateSheet_(ss, SHEET_SETTINGS);
  const keywordsSheet = getOrCreateSheet_(ss, SHEET_KEYWORDS);
  ensureSettingRows_();

  const keywords = readKeywords_(keywordsSheet);

  const input = {
    queries: keywords,
    countryCode: String(readSetting_(settings, 'countryCode') || 'us').toLowerCase(),
    maxResultsPerQuery: Number(readSetting_(settings, 'maxResultsPerQuery') || 10),
    maxPagesPerQuery: Number(readSetting_(settings, 'maxPagesPerQuery') || 2),
    outputFormat: String(readSetting_(settings, 'outputFormat') || 'flat'),
    mobileResults: parseBool_(readSetting_(settings, 'mobileResults')),
    disableGoogleSearchResults: parseBool_(readSetting_(settings, 'disableGoogleSearchResults')),
    forceExactMatch: parseBool_(readSetting_(settings, 'forceExactMatch')),
    focusOnPaidAds: parseBool_(readSetting_(settings, 'focusOnPaidAds')),
    includeUnfilteredResults: parseBool_(readSetting_(settings, 'includeUnfilteredResults')),
    saveHtml: parseBool_(readSetting_(settings, 'saveHtml')),
    saveHtmlToKeyValueStore: parseBool_(readSetting_(settings, 'saveHtmlToKeyValueStore')),
    includeIcons: parseBool_(readSetting_(settings, 'includeIcons')),
    llmSearchScope: String(readSetting_(settings, 'llmSearchScope') || 'perQuery'),
    llmQueryMode: String(readSetting_(settings, 'llmQueryMode') || 'sameAsKeyword'),
    llmApplyTo: String(readSetting_(settings, 'llmApplyTo') || 'brandLike'),
    chatGptSearch: { enableChatGpt: parseBool_(readSetting_(settings, 'enableChatGpt')) },
    geminiSearch: { enableGemini: parseBool_(readSetting_(settings, 'enableGemini')) },
    copilotSearch: { enableCopilot: parseBool_(readSetting_(settings, 'enableCopilot')) },
    deepSeekSearch: { enableDeepSeek: parseBool_(readSetting_(settings, 'enableDeepSeek')) },
    aiModeSearch: { enableAiMode: parseBool_(readSetting_(settings, 'enableAiMode')) },
    perplexitySearch: {
      enablePerplexity: parseBool_(readSetting_(settings, 'enablePerplexity')),
      searchRecency: String(readSetting_(settings, 'perplexitySearchRecency') || 'month'),
      returnImages: parseBool_(readSetting_(settings, 'perplexityReturnImages')),
      returnRelatedQuestions: parseBool_(readSetting_(settings, 'perplexityReturnRelatedQuestions')),
    },
  };

  assignOptionalNumber_(input, 'maxTotalResults', readOptionalNumber_(settings, 'maxTotalResults'));
  assignOptionalString_(input, 'site', readOptionalString_(settings, 'site'));
  assignOptionalString_(input, 'relatedToSite', readOptionalString_(settings, 'relatedToSite'));
  assignOptionalString_(input, 'afterDate', readOptionalString_(settings, 'afterDate'));
  assignOptionalString_(input, 'beforeDate', readOptionalString_(settings, 'beforeDate'));
  assignOptionalString_(input, 'searchLanguage', readOptionalString_(settings, 'searchLanguage'));
  assignOptionalString_(input, 'languageCode', readOptionalString_(settings, 'languageCode'));
  assignOptionalString_(input, 'locationUule', readOptionalString_(settings, 'locationUule'));
  assignOptionalString_(input, 'targetBrand', readOptionalString_(settings, 'targetBrand'));
  assignOptionalString_(input, 'quickDateRange', readOptionalQuickDate_(settings, 'quickDateRange'));
  assignOptionalStringList_(input, 'wordsInTitle', readStringList_(settings, 'wordsInTitle'));
  assignOptionalStringList_(input, 'wordsInText', readStringList_(settings, 'wordsInText'));
  assignOptionalStringList_(input, 'wordsInUrl', readStringList_(settings, 'wordsInUrl'));
  assignOptionalStringList_(input, 'fileTypes', readStringList_(settings, 'fileTypes'));
  assignOptionalStringList_(input, 'targetDomains', readStringList_(settings, 'targetDomains'));

  return input;
}

/** Keywords from column A (row 2+). Uses A2:A so getLastRow() gaps do not hide data. */
function readKeywords_(keywordsSheet) {
  const values = keywordsSheet.getRange('A2:A').getValues();
  return values
    .map(function (row) {
      return String(row[0] || '').trim();
    })
    .filter(Boolean);
}

function ensureSettingRows_() {
  if (SETTING_ROW.status) {
    return;
  }
  rebuildSettingRows_();
}

function readSetting_(sheet, key) {
  ensureSettingRows_();
  const row = SETTING_ROW[key];
  if (!row) {
    throw new Error('Unknown setting key: ' + key);
  }
  return sheet.getRange(row, 2).getValue();
}

function readOptionalString_(sheet, key) {
  const value = String(readSetting_(sheet, key) || '').trim();
  return value || undefined;
}

function readOptionalNumber_(sheet, key) {
  const raw = readSetting_(sheet, key);
  if (raw === '' || raw === null || raw === undefined) {
    return undefined;
  }
  const normalized = String(raw).trim();
  if (normalized === '(none)') {
    return undefined;
  }
  const number = Number(normalized);
  return isNaN(number) ? undefined : number;
}

function readOptionalQuickDate_(sheet, key) {
  const value = String(readSetting_(sheet, key) || '').trim();
  if (!value || value === '(none)') {
    return undefined;
  }
  return value;
}

function readStringList_(sheet, key) {
  const value = String(readSetting_(sheet, key) || '').trim();
  if (!value) {
    return undefined;
  }
  const parts = value
    .split(/[,\n]/)
    .map(function (part) {
      return part.trim();
    })
    .filter(Boolean);
  return parts.length ? parts : undefined;
}

function assignOptionalString_(target, key, value) {
  if (value) {
    target[key] = value;
  }
}

function assignOptionalNumber_(target, key, value) {
  if (value != null) {
    target[key] = value;
  }
}

function assignOptionalStringList_(target, key, value) {
  if (value && value.length) {
    target[key] = value;
  }
}

function parseBool_(value) {
  const normalized = String(value).trim().toLowerCase();
  return normalized === 'true' || normalized === 'yes' || normalized === '1';
}

// ─── Dataset import ───────────────────────────────────────────────────────

function writeAllResults_(token, datasetId) {
  const items = fetchDatasetItems_(token, datasetId);
  const flatRows = items.filter(function (row) {
    return row.recordType === 'flat';
  });
  let llmRows = items.filter(isLlmDatasetRow_);
  let llmSource = 'dataset';
  if (llmRows.length === 0) {
    llmRows = buildLlmRowsFromFlat_(flatRows);
    if (llmRows.length > 0) {
      llmSource = 'flat-fallback';
    }
  }
  const usageRow = items.find(function (row) {
    return row.recordType === 'usage_summary';
  });

  writeFlatResults_(flatRows);
  const llmAnswers = writeLlmAnswers_(llmRows);
  const llmCitations = writeLlmCitations_(llmRows);
  writeLlmSummary_(llmRows);

  return {
    flat: flatRows.length,
    llm: llmRows.length,
    llmAnswers: llmAnswers,
    llmCitations: llmCitations,
    llmSource: llmSource,
    usage: usageRow || null,
  };
}

function isLlmDatasetRow_(row) {
  if (row.recordType === 'llm') {
    return true;
  }
  return Boolean(
    row.chatGptSearchResult ||
      row.geminiSearchResult ||
      row.perplexitySearchResult ||
      row.copilotSearchResult ||
      row.deepSeekSearchResult,
  );
}

/** Build per-keyword LLM summary when dataset has no recordType=llm rows (older Actor builds). */
function buildLlmRowsFromFlat_(flatRows) {
  const hasLlmColumns = flatRows.some(function (row) {
    return row.llmChatGptCited || row.llmGeminiCited || row.llmPerplexityCited;
  });
  if (!hasLlmColumns) {
    return [];
  }

  const groups = {};
  flatRows.forEach(function (row) {
    const keyword = String(row.keyword || '').trim();
    if (!keyword) {
      return;
    }
    if (!groups[keyword]) {
      groups[keyword] = {
        keyword: keyword,
        targetCountry: row.targetCountry || '',
        fetchedAt: row.fetchedAt || '',
        googleUrls: [],
        chatGptUrls: [],
        geminiUrls: [],
        perplexityUrls: [],
      };
    }
    const group = groups[keyword];
    if (row.url) {
      group.googleUrls.push(row.url);
    }
    if (row.llmChatGptCited && row.url) {
      group.chatGptUrls.push(row.url);
    }
    if (row.llmGeminiCited && row.url) {
      group.geminiUrls.push(row.url);
    }
    if (row.llmPerplexityCited && row.url) {
      group.perplexityUrls.push(row.url);
    }
  });

  return Object.keys(groups).map(function (keyword) {
    const group = groups[keyword];
    const chatGptUrls = uniqueUrls_(group.chatGptUrls);
    const geminiUrls = uniqueUrls_(group.geminiUrls);
    const perplexityUrls = uniqueUrls_(group.perplexityUrls);
    const allLlmUrls = uniqueUrls_(chatGptUrls.concat(geminiUrls).concat(perplexityUrls));
    const googleTop = uniqueUrls_(group.googleUrls).slice(0, 10);
    const primaryDomain = domainFromUrl_(googleTop[0] || '');
    const googleRank = primaryDomain ? rankForDomain_(googleTop, primaryDomain) : '';
    const llmRank = primaryDomain ? rankForDomain_(allLlmUrls, primaryDomain) : '';

    return {
      keyword: group.keyword,
      targetCountry: group.targetCountry,
      fetchedAt: group.fetchedAt,
      queryPage: 1,
      chatGptSearchResult: {
        query: group.keyword,
        citedUrls: chatGptUrls,
        citations: urlsToCitations_(chatGptUrls),
      },
      geminiSearchResult: {
        query: group.keyword,
        citedUrls: geminiUrls,
        citations: urlsToCitations_(geminiUrls),
      },
      perplexitySearchResult: {
        query: group.keyword,
        citedUrls: perplexityUrls,
        citations: urlsToCitations_(perplexityUrls),
      },
      visibilityCompare: {
        overlapPercent: overlapPercentFromUrls_(googleTop, allLlmUrls),
        primaryDomain: primaryDomain,
        targetDomainRanks: primaryDomain
          ? [
              {
                googleRank: googleRank,
                llmRank: llmRank,
                rankDelta: googleRank && llmRank ? llmRank - googleRank : '',
              },
            ]
          : [],
      },
    };
  });
}

function uniqueUrls_(urls) {
  const seen = {};
  const out = [];
  urls.forEach(function (url) {
    const normalized = String(url || '').trim();
    if (!normalized || seen[normalized]) {
      return;
    }
    seen[normalized] = true;
    out.push(normalized);
  });
  return out;
}

function domainFromUrl_(url) {
  const raw = String(url || '').trim();
  if (!raw) {
    return '';
  }
  const withoutProtocol = raw.replace(/^https?:\/\//i, '');
  const host = withoutProtocol.split('/')[0].toLowerCase();
  return host.replace(/^www\./, '');
}

function overlapPercentFromUrls_(googleUrls, llmUrls) {
  const googleDomains = googleUrls.map(domainFromUrl_).filter(Boolean);
  const llmDomains = llmUrls.map(domainFromUrl_).filter(Boolean);
  if (!googleDomains.length || !llmDomains.length) {
    return '';
  }
  const googleSet = {};
  googleDomains.forEach(function (domain) {
    googleSet[domain] = true;
  });
  let matches = 0;
  llmDomains.forEach(function (domain) {
    if (googleSet[domain]) {
      matches++;
    }
  });
  return Math.round((matches / Math.max(googleDomains.length, llmDomains.length)) * 100);
}

function rankForDomain_(urls, domain) {
  for (let i = 0; i < urls.length; i++) {
    if (domainFromUrl_(urls[i]) === domain) {
      return i + 1;
    }
  }
  return '';
}

function writeFlatResults_(flatRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_RESULTS);
  sheet.clear();
  writeHeaderRow_(sheet, FLAT_COLUMNS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, FLAT_COLUMNS.length).setBackground('#e8f0fe');

  if (flatRows.length === 0) {
    return;
  }

  const values = flatRows.map(function (row) {
    return FLAT_COLUMNS.map(function (col) {
      const val = row[col];
      if (val === null || val === undefined) {
        return '';
      }
      return val;
    });
  });

  sheet.getRange(2, 1, values.length, FLAT_COLUMNS.length).setValues(values);
  sheet.autoResizeColumns(1, FLAT_COLUMNS.length);
}

function writeLlmSummary_(llmRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_LLM);
  const columns = getLlmSummaryColumns_();
  sheet.clear();
  writeHeaderRow_(sheet, columns);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, columns.length).setBackground('#e6f4ea');

  if (llmRows.length === 0) {
    return;
  }

  const values = llmRows.map(function (row) {
    const compare = row.visibilityCompare || {};
    const target = (compare.targetDomainRanks && compare.targetDomainRanks[0]) || {};
    const out = [
      row.keyword || '',
      row.targetCountry || '',
      row.queryPage != null ? row.queryPage : '',
      row.fetchedAt || '',
      compare.overlapPercent != null ? compare.overlapPercent : '',
      compare.primaryDomain || '',
      target.googleRank != null ? target.googleRank : '',
      target.llmRank != null ? target.llmRank : '',
      target.rankDelta != null ? target.rankDelta : '',
    ];
    LLM_PROVIDER_EXPORTS.forEach(function (provider) {
      const result = row[provider.resultKey] || {};
      out.push(
        result.query || '',
        result.webSearchQuery || '',
        citationCount_(result),
        result.error || '',
      );
    });
    return out;
  });

  sheet.getRange(2, 1, values.length, columns.length).setValues(values);
  sheet.autoResizeColumns(1, columns.length);
}

/** One row per keyword × provider — full answer text for analytics. */
function writeLlmAnswers_(llmRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_LLM_ANSWERS);
  sheet.clear();
  writeHeaderRow_(sheet, LLM_ANSWERS_COLUMNS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, LLM_ANSWERS_COLUMNS.length).setBackground('#d9ead3');

  const values = [];
  llmRows.forEach(function (row) {
    LLM_PROVIDER_EXPORTS.forEach(function (provider) {
      const result = row[provider.resultKey];
      if (!hasLlmProviderData_(result)) {
        return;
      }
      values.push([
        row.keyword || '',
        row.targetCountry || '',
        row.queryPage != null ? row.queryPage : '',
        row.fetchedAt || '',
        provider.label,
        result.query || '',
        result.webSearchQuery || '',
        result.model || '',
        formatLlmAnswer_(result),
        citationCount_(result),
        result.error || '',
      ]);
    });
  });

  if (values.length === 0) {
    return 0;
  }

  sheet.getRange(2, 1, values.length, LLM_ANSWERS_COLUMNS.length).setValues(values);
  sheet.getRange(2, 9, values.length, 1).setWrap(true);
  sheet.autoResizeColumns(1, LLM_ANSWERS_COLUMNS.length);
  return values.length;
}

/** One row per citation (matches dataset JSON citations[]). */
function writeLlmCitations_(llmRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_LLM_CITATIONS);
  sheet.clear();
  writeHeaderRow_(sheet, LLM_CITATIONS_COLUMNS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, LLM_CITATIONS_COLUMNS.length).setBackground('#fff2cc');

  const values = [];
  llmRows.forEach(function (row) {
    LLM_PROVIDER_EXPORTS.forEach(function (provider) {
      const result = row[provider.resultKey];
      if (!result) {
        return;
      }
      normalizeCitations_(result).forEach(function (citation) {
        values.push([
          row.keyword || '',
          row.targetCountry || '',
          row.queryPage != null ? row.queryPage : '',
          row.fetchedAt || '',
          provider.label,
          result.query || '',
          result.webSearchQuery || '',
          citation.rank != null ? citation.rank : '',
          citation.url || '',
          citation.title || '',
        ]);
      });
    });
  });

  if (values.length === 0) {
    return 0;
  }

  sheet.getRange(2, 1, values.length, LLM_CITATIONS_COLUMNS.length).setValues(values);
  sheet.autoResizeColumns(1, LLM_CITATIONS_COLUMNS.length);
  return values.length;
}

function hasLlmProviderData_(result) {
  if (!result) {
    return false;
  }
  return Boolean(
    result.answer ||
      result.error ||
      citationCount_(result) > 0 ||
      result.query ||
      result.webSearchQuery,
  );
}

function citationCount_(result) {
  if (!result) {
    return 0;
  }
  if (result.citations && result.citations.length) {
    return result.citations.length;
  }
  if (result.citedUrls && result.citedUrls.length) {
    return result.citedUrls.length;
  }
  return 0;
}

function normalizeCitations_(result) {
  if (!result) {
    return [];
  }
  if (result.citations && result.citations.length) {
    return result.citations.map(function (citation) {
      return {
        rank: citation.rank != null ? citation.rank : '',
        url: citation.url || '',
        title: citation.title || '',
      };
    });
  }
  return (result.citedUrls || []).map(function (url, index) {
    return { rank: index + 1, url: url, title: '' };
  });
}

function urlsToCitations_(urls) {
  return (urls || []).map(function (url, index) {
    return { rank: index + 1, url: url, title: '' };
  });
}

function formatLlmAnswer_(result) {
  if (!result) {
    return '';
  }
  if (result.answer) {
    return cellText_(result.answer, LLM_CELL_CHAR_LIMIT);
  }
  if (result.error) {
    return '[error] ' + result.error;
  }
  return '';
}

function cellText_(value, maxLen) {
  const text = value == null ? '' : String(value);
  if (!maxLen || text.length <= maxLen) {
    return text;
  }
  return text.slice(0, maxLen - 15) + '…[truncated]';
}

function appendRunLog_(runId, status, counts, note, run) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_RUN_LOG);
  if (sheet.getLastRow() === 0) {
    setupRunLogSheet_(ss);
  } else {
    syncRunLogHeader_(sheet);
  }

  const usageRow = counts && counts.usage ? counts.usage : null;
  const billing = usageRow && usageRow.billing ? usageRow.billing : null;
  const llmCalls = billing && billing.llmCallsCharged ? billing.llmCallsCharged : {};

  const row = [
    new Date().toISOString(),
    runId,
    status,
    billing ? billing.organicUrlsCharged : '',
    llmCalls.chatGpt != null ? llmCalls.chatGpt : '',
    llmCalls.gemini != null ? llmCalls.gemini : '',
    llmCalls.perplexity != null ? llmCalls.perplexity : '',
    llmCalls.deepSeek != null ? llmCalls.deepSeek : '',
    billing ? billing.pagesFetched : '',
    formatRunCostUsd_(run),
    note || '',
  ];

  sheet.appendRow(row);
}

/** Upgrade header when costUsd column was added after an older template version. */
function syncRunLogHeader_(sheet) {
  if (sheet.getRange(1, 10).getValue() === 'costUsd') {
    return;
  }
  writeHeaderRow_(sheet, RUN_LOG_COLUMNS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, RUN_LOG_COLUMNS.length).setBackground('#fce8e6');
}

/** Apify run.usageTotalUsd — actual USD charged (requires API token). */
function formatRunCostUsd_(run) {
  if (!run || run.usageTotalUsd == null || run.usageTotalUsd === '') {
    return '';
  }
  const amount = Number(run.usageTotalUsd);
  return isNaN(amount) ? '' : amount;
}

function fetchDatasetItems_(token, datasetId) {
  const url =
    'https://api.apify.com/v2/datasets/' +
    datasetId +
    '/items?format=json&clean=false&limit=10000';
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() >= 300) {
    throw new Error('Dataset fetch failed: ' + response.getContentText().slice(0, 200));
  }

  return JSON.parse(response.getContentText());
}

// ─── Helpers ──────────────────────────────────────────────────────────────

function setStatusNote_(message) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
    if (sheet) {
      ensureSettingRows_();
      sheet.getRange(SETTING_ROW.status, 2).setValue(message);
    }
  } catch (e) {
    // ignore if trigger context has no active spreadsheet focus
  }
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function startActorRun_(token, actorId, input) {
  if (!Array.isArray(input.queries) || input.queries.length === 0) {
    throw new Error('Actor input has no queries. Check the Keywords sheet (column A).');
  }
  const encodedActor = encodeURIComponent(actorId.replace('/', '~'));
  // POST body is the Actor input directly (Apify /v2/actors/.../runs).
  return apifyPost_(token, '/v2/actors/' + encodedActor + '/runs', input);
}

function apifyGet_(token, path) {
  const response = UrlFetchApp.fetch('https://api.apify.com' + path, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() >= 300) {
    throw new Error('Apify GET ' + path + ': ' + response.getContentText().slice(0, 200));
  }
  const body = JSON.parse(response.getContentText());
  return body.data || body;
}

function apifyPost_(token, path, payload) {
  const response = UrlFetchApp.fetch('https://api.apify.com' + path, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  const body = JSON.parse(response.getContentText());
  if (response.getResponseCode() >= 300) {
    throw new Error('Apify API error: ' + response.getContentText().slice(0, 300));
  }
  return body.data;
}
