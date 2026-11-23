/**
 * Format QA Issue Report v2  –  Google Apps Script (bound to a Google Doc)
 *
 * WHAT IT DOES
 *  - Dark-blue title banner (white text) with issue count + date
 *  - Soft page background colour
 *  - Screenshots scaled to fit, side-by-side screenshot tables turned into white "cards"
 *  - Clean numbered issues: bold headline sentence, consistent spacing, clickable links
 *  - Native Google Docs status dropdown chips beside each numbered issue
 *  - Temporary status-filtered tabs without changing the original report
 *  - Header + footer, tighter margins, one font everywhere. Wording is never changed.
 *
 * HOW TO USE
 *  1. Upload the .docx to Drive > right-click > Open with > Google Docs.
 *  2. In the Doc: Extensions > Apps Script, delete old code, paste this, Save.
 *  3. Run `formatQAReport`, authorise the requested permissions, then reload the Doc.
 *     It adds the native status chips automatically; use the Format report menu to
 *     add them later or update existing chips without reformatting the report.
 *  Tip: run it on a copy first (File > Make a copy).
 */

var CFG = {
  FONT: 'Arial',
  BODY_PT: 11,
  TEXT_COLOR: '#1F2933',
  HEADLINE_COLOR: '#111111',      // bold first line of each issue (independent of banner colour)
  ACCENT: '#1F3A5F',
  BANNER_BG: '#1F3A5F',
  BANNER_TITLE_PT: 24,
  BANNER_SUB_COLOR: '#BFD3F2',
  PAGE_BG: '#F4F6FA',             // soft blue-grey page background
  CARD_BG: '#FFFFFF',
  CARD_BORDER: '#D0D7E2',
  LINK_COLOR: '#1A73E8',
  // bg = chip background, fg = chip text colour (what you pick when editing a dropdown in Docs)
  STATUS_OPTIONS: [
    { id: 'dropdownItem.open',     value: 'Open',      bg: '#D93025', fg: '#FFFFFF' },
    { id: 'dropdownItem.hold',     value: 'Hold',      bg: '#F9AB00', fg: '#202124' },
    { id: 'dropdownItem.notfixed', value: 'Not fixed', bg: '#E8710A', fg: '#FFFFFF' },
    { id: 'dropdownItem.notbug',   value: 'Not a bug', bg: '#80868B', fg: '#FFFFFF' },
    { id: 'dropdownItem.fixed',    value: 'Fixed',     bg: '#188038', fg: '#FFFFFF' }
  ],
  MARGIN_PT: 54,
  PAGE_W: 612, PAGE_H: 792,       // US Letter. A4 = 595.28 x 841.89
  FIT_IMAGES: false,            // set from the options window: true = shrink images wider than the page
  IMG_MAX_W: 468,               // images are only shrunk if wider than this; never enlarged, no height cap
  IMG_MAX_W_IN_TABLE: 225,      // images inside the 2-column tables must fit their cell
  ISSUE_SPACE_BEFORE: 10,
  ISSUE_SPACE_AFTER: 6,
  HEADER_TEXT: 'QA Issue Report – Admin User POS',
  FOOTER_TEXT: 'Internal – QA findings'   // page numbers: Insert > Page numbers
};

/** Adds a "Format report" menu to the Doc (appears after you reload it). */
function onOpen() {
  var ui = DocumentApp.getUi();
  var filterMenu = ui.createMenu('Filter Issues')
    .addItem('Show All Issues', 'showAllIssues')
    .addSeparator()
    .addItem('Open', 'filterOpenIssues')
    .addItem('Hold', 'filterHoldIssues')
    .addItem('Fixed', 'filterFixedIssues')
    .addItem('Not a Bug', 'filterNotABugIssues')
    .addItem('Not fixed', 'filterNotFixedIssues')
    .addSeparator()
    .addItem('Restore Original View', 'restoreOriginalIssueView');

  ui.createMenu('Format report')
    .addItem('Format QA report…', 'formatQAReport')
    .addItem('Add issue status dropdowns', 'addIssueStatusDropdowns')
    .addSeparator()
    .addSubMenu(filterMenu)
    .addToUi();
}

function filterOpenIssues() {
  filterIssuesByStatus_('Open');
}

function filterHoldIssues() {
  filterIssuesByStatus_('Hold');
}

function filterFixedIssues() {
  filterIssuesByStatus_('Fixed');
}

function filterNotABugIssues() {
  filterIssuesByStatus_('Not a bug');
}

function filterNotFixedIssues() {
  filterIssuesByStatus_('Not fixed');
}

function showAllIssues() {
  showOriginalIssueDocument_(false);
}

function restoreOriginalIssueView() {
  showOriginalIssueDocument_(true);
}

/**
 * ENTRY POINT – run this (or use the Doc menu: Format report > Format QA report).
 * It opens an options window INSIDE THE GOOGLE DOC. The Apps Script editor only shows
 * "Execution completed" – that is normal; switch to the Doc tab and fill in the options.
 */
function formatQAReport() {
  try {
    var html = HtmlService.createHtmlOutput(buildDialogHtml_()).setWidth(460).setHeight(740);
    DocumentApp.getUi().showModalDialog(html, 'Format QA report');
    Logger.log('>>> A window has opened in your Google Doc. Go to the DOC tab, choose your options and click "Format document". <<<');
  } catch (e) {
    Logger.log('Could not open the options window from here: ' + e + '\n' +
      'FIX: keep the Google Doc open in another browser tab, reload it, then use the Doc menu: Format report > Format QA report.\n' +
      'Or run formatWithDefaults() to format right now without the window.');
  }
}

/** Runs the formatting straight from the editor with the default choices (no window). */
function formatWithDefaults() {
  var r = applyFormatting({ listType: 'NUMBER', dividers: true, fitImages: false, pageBg: true, bgColor: CFG.PAGE_BG, accent: CFG.ACCENT, makeOutline: true });
  Logger.log(r);
}

/** Called by the dialog when the user clicks "Format document". */
function applyFormatting(opts) {
  if (!opts) {                       // someone pressed Run on this function directly
    formatQAReport();
    return 'opened-dialog';
  }
  var hex = /^#[0-9a-fA-F]{6}$/;
  var glyphName = ['BULLET', 'HOLLOW_BULLET', 'SQUARE_BULLET', 'NUMBER', 'LATIN_UPPER', 'LATIN_LOWER', 'ROMAN_UPPER', 'ROMAN_LOWER']
                    .indexOf(opts.listType) > -1 ? opts.listType : 'NUMBER';
  if (hex.test(opts.accent || '')) { CFG.ACCENT = opts.accent; CFG.BANNER_BG = opts.accent; }
  var bgColor = hex.test(opts.bgColor || '') ? opts.bgColor : CFG.PAGE_BG;
  CFG.GLYPH = DocumentApp.GlyphType[glyphName];
  CFG.FIT_IMAGES = !!opts.fitImages;

  var doc = DocumentApp.getActiveDocument();
  var docId = doc.getId();
  var body = doc.getBody();
  var failed = [];
  function step(name, fn) {          // one failing step must not stop the others
    try { fn(); Logger.log('OK      ' + name); }
    catch (e) {
      var message = e && e.message ? e.message : String(e);
      failed.push(name + ' (' + message + ')');
      Logger.log('FAILED  ' + name + ' -> ' + message);
    }
  }

  step('status outline cleanup', function () {
    normalizeIssueStatusDropdownParagraphs_(docId);
  });
  var issueCount = countIssues_(body);
  step('page + base style',   function () { setPageAndBaseStyle_(body); });
  step('tidy blank lines',    function () { tidyEmptyParagraphs_(body); });
  step('style issues/images', function () { styleIssues_(body); });
  if (opts.dividers) step('separator lines', function () { addIssueDividers_(body); });
  step('screenshot tables',   function () { styleScreenshotTables_(body); });
  step('title banner',        function () { buildBanner_(body, issueCount); });
  if (opts.makeOutline) step('document outline', function () { createIssueOutline_(body, glyphName); });
  step('links',               function () { linkifyUrls_(body); });
  step('header + footer',     function () { addHeaderFooter_(doc); });

  doc.saveAndClose();
  if (opts.makeOutline) {
    step('outline headings', function () { applyIssueOutlineStyles_(docId); });
  }
  step('issue status dropdowns', function () { addIssueStatusDropdowns_(docId); });
  if (opts.pageBg) setPageBackground_(docId, bgColor);
  return failed.length ? ('Finished, but these steps failed: ' + failed.join(', ') + '. See View > Logs in the script editor.') : 'done';
}

/** Adds native Google Docs dropdown chips beside numbered issues. */
function addIssueStatusDropdowns() {
  var result = addIssueStatusDropdowns_(DocumentApp.getActiveDocument().getId());
  DocumentApp.getUi().alert(result);
}

function addIssueStatusDropdowns_(docId) {
  var api = 'https://docs.googleapis.com/v1/documents/' + encodeURIComponent(docId);
  var token = ScriptApp.getOAuthToken();
  var options = CFG.STATUS_OPTIONS.map(function (option) {
    return {
      optionId: option.id,
      displayValue: option.value,
      textStyle: {
        foregroundColor: { color: { rgbColor: hexToRgb_(option.fg || '#FFFFFF') } },
        backgroundColor: { color: { rgbColor: hexToRgb_(option.bg || '#FFFFFF') } }
      }
    };
  });
  var desiredByName = {};
  options.forEach(function (option) {
    desiredByName[option.displayValue] = option.optionId;
  });

  function readDocState() {
    var doc = docsApiRequest_(api + '?includeTabsContent=true', 'get', token);
    var content = doc.body && doc.body.content || [];
    var definitionMap = {};
    var tabId = null;
    if (doc.tabs && doc.tabs.length && doc.tabs[0].documentTab) {
      var mainTab = doc.tabs[0];
      var tab = mainTab.documentTab;
      if (tab.body) content = tab.body.content || content;
      definitionMap = tab.dropdownDefinitions || {};
      tabId = mainTab.tabProperties && mainTab.tabProperties.tabId || null;
    }
    return { content: content, definitions: definitionMap, tabId: tabId };
  }

  var state = readDocState();
  var definitions = state.definitions;
  var definitionMap = {};
  Object.keys(definitions).forEach(function (key) {
    var definition = definitions[key];
    definitionMap[definition.dropdownDefinitionId] = definition;
  });

  var issueElements = [];
  var deleteRequests = [];
  var issueStatuses = [];
  var statusLineExists = [];
  state.content.forEach(function (item, itemIndex) {
    if (!isApiIssueHeading_(item)) return;
    issueElements.push(item);

    var dropdown = (item.paragraph.elements || []).filter(function (element) {
      return !!element.dropdown;
    })[0];
    var followingItem = state.content[itemIndex + 1];
    var followingDropdown = null;
    if (!dropdown && followingItem && followingItem.paragraph &&
        !isApiIssueHeading_(followingItem)) {
      followingDropdown = (followingItem.paragraph.elements || []).filter(function (element) {
        return !!element.dropdown;
      })[0] || null;
    }
    var status = options[0].displayValue;
    var existingDropdown = dropdown || followingDropdown;
    if (existingDropdown) {
      var props = existingDropdown.dropdown.dropdownProperties || {};
      var definition = definitionMap[props.dropdownDefinitionId];
      var definitionOptions = definition &&
        definition.dropdownDefinitionProperties &&
        definition.dropdownDefinitionProperties.options || [];
      for (var i = 0; i < definitionOptions.length; i++) {
        if (definitionOptions[i].optionId === props.selectedOptionId) {
          if (desiredByName[definitionOptions[i].displayValue]) {
            status = definitionOptions[i].displayValue;
          }
          break;
        }
      }
    }
    if (dropdown) {
      deleteRequests.push({
        deleteContentRange: {
          range: {
            startIndex: dropdown.startIndex,
            endIndex: dropdown.endIndex
          }
        }
      });
    }
    issueStatuses.push(status);
    statusLineExists.push(!!followingDropdown);
  });

  if (!issueElements.length) return 'No numbered issue list items were found.';

  deleteRequests.sort(function (a, b) {
    return b.deleteContentRange.range.startIndex -
      a.deleteContentRange.range.startIndex;
  });
  if (deleteRequests.length) {
    docsApiRequest_(api + ':batchUpdate', 'post', token, {
      requests: deleteRequests
    });
  }

  state = readDocState();
  var splitImageRequests = [];
  var splitImageIndexes = [];
  state.content.forEach(function (item) {
    if (!isApiIssueHeading_(item)) return;
    var elements = item.paragraph.elements || [];
    var image = null;
    var hasTextBeforeImage = false;
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      if (element.inlineObjectElement) {
        image = element;
        break;
      }
      if (element.textRun && /\S/.test(element.textRun.content)) {
        hasTextBeforeImage = true;
      }
    }
    if (image && hasTextBeforeImage) {
      splitImageIndexes.push(image.startIndex);
      splitImageRequests.push({
        insertText: {
          location: { index: image.startIndex },
          text: '\n'
        }
      });
    }
  });
  splitImageRequests.sort(function (a, b) {
    return b.insertText.location.index - a.insertText.location.index;
  });
  if (splitImageRequests.length) {
    docsApiRequest_(api + ':batchUpdate', 'post', token, {
      requests: splitImageRequests
    });

    state = readDocState();
    var continuationBulletRanges = [];
    var continuationStarts = {};
    splitImageIndexes.forEach(function (index) {
      var precedingSplits = splitImageIndexes.filter(function (otherIndex) {
        return otherIndex < index;
      }).length;
      continuationStarts[index + 1 + precedingSplits] = true;
    });
    state.content.forEach(function (item) {
      if (!item.paragraph || !item.paragraph.bullet) return;
      if (continuationStarts[item.startIndex]) {
        continuationBulletRanges.push({
          startIndex: item.startIndex,
          endIndex: item.endIndex
        });
      }
    });
    continuationBulletRanges.sort(function (a, b) {
      return b.startIndex - a.startIndex;
    });
    if (continuationBulletRanges.length) {
      docsApiRequest_(api + ':batchUpdate', 'post', token, {
        requests: continuationBulletRanges.map(function (range) {
          return { deleteParagraphBullets: { range: range } };
        })
      });
    }
  }

  var dropdownDefinitionId = null;
  Object.keys(definitions).some(function (key) {
    var definition = definitions[key];
    var definitionOptions = definition.dropdownDefinitionProperties &&
      definition.dropdownDefinitionProperties.options || [];
    var matches = definitionOptions.length === options.length;
    for (var i = 0; matches && i < options.length; i++) {
      matches = definitionOptions[i].optionId === options[i].optionId &&
        definitionOptions[i].displayValue === options[i].displayValue &&
        styleKey_(definitionOptions[i].textStyle) === styleKey_(options[i].textStyle);
    }
    if (matches) {
      dropdownDefinitionId = definition.dropdownDefinitionId;
      return true;
    }
    return false;
  });

  if (!dropdownDefinitionId) {
    var createResponse = docsApiRequest_(api + ':batchUpdate', 'post', token, {
      requests: [{
        createDropdownDefinition: {
          dropdownDefinition: {
            dropdownDefinitionProperties: {
              title: 'Issue status',
              options: options
            }
          }
        }
      }]
    });
    var created = createResponse.replies && createResponse.replies[0] &&
      createResponse.replies[0].createDropdownDefinition &&
      createResponse.replies[0].createDropdownDefinition.dropdownDefinition;
    dropdownDefinitionId = created && created.dropdownDefinitionId;
    if (!dropdownDefinitionId) {
      throw new Error('Google Docs did not return the new dropdown definition ID.');
    }
  }

  state = readDocState();
  var insertionPlans = [];
  var statusIndex = 0;
  state.content.forEach(function (item) {
    if (!isApiIssueHeading_(item)) return;
    if (!statusLineExists[statusIndex]) {
      insertionPlans.push({
        index: item.endIndex - 1,
        statusOptionId: desiredByName[issueStatuses[statusIndex]]
      });
    }
    statusIndex++;
  });

  insertionPlans.sort(function (a, b) { return b.index - a.index; });
  var insertRequests = [];
  insertionPlans.forEach(function (plan) {
    var statusIndex = plan.index + 1;
    insertRequests.push({
      insertText: { location: { index: plan.index }, text: '\n' }
    });
    insertRequests.push({
      deleteParagraphBullets: {
        range: { startIndex: statusIndex, endIndex: statusIndex + 1 }
      }
    });
    insertRequests.push({
      insertDropdown: {
        location: { index: statusIndex },
        dropdownDefinitionId: dropdownDefinitionId,
        selectedOptionId: plan.statusOptionId
      }
    });
  });

  if (!insertRequests.length) {
    normalizeIssueStatusDropdownParagraphs_(docId);
    return 'Status dropdowns are already present for all ' +
      issueElements.length + ' issue(s).';
  }

  docsApiRequest_(api + ':batchUpdate', 'post', token, {
    requests: insertRequests
  });
  normalizeIssueStatusDropdownParagraphs_(docId);
  return 'Positioned status dropdowns beneath the issue text for ' +
    issueElements.length + ' issue(s).';
}

function normalizeIssueStatusDropdownParagraphs_(docId) {
  var api = 'https://docs.googleapis.com/v1/documents/' + encodeURIComponent(docId);
  var document = docsApiRequest_(
    api + '?includeTabsContent=true',
    'get',
    ScriptApp.getOAuthToken()
  );
  var mainTab = document.tabs && document.tabs[0];
  var content = mainTab && mainTab.documentTab && mainTab.documentTab.body
    ? mainTab.documentTab.body.content || []
    : document.body && document.body.content || [];
  var tabId = mainTab && mainTab.tabProperties && mainTab.tabProperties.tabId;
  var requests = [];
  content.forEach(function (item) {
    if (!item.paragraph || isApiIssueHeading_(item) ||
        !findDropdownInParagraph_(item.paragraph)) return;
    var range = { startIndex: item.startIndex, endIndex: item.endIndex };
    if (tabId) range.tabId = tabId;
    if (item.paragraph.bullet) {
      requests.push({ deleteParagraphBullets: { range: range } });
    }
    requests.push({
      updateParagraphStyle: {
        range: range,
        paragraphStyle: { namedStyleType: 'NORMAL_TEXT' },
        fields: 'namedStyleType'
      }
    });
  });
  if (requests.length) {
    docsApiRequest_(api + ':batchUpdate', 'post', ScriptApp.getOAuthToken(), {
      requests: requests
    });
  }
}

function docsApiRequest_(url, method, token, payload) {
  var options = {
    method: method,
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  };
  if (payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Google Docs API error (' + code + '): ' + body);
  }
  return body ? JSON.parse(body) : {};
}

var FILTER_TAB_PROPERTY_ = 'FORMAT_QA_FILTER_TAB_ID';
var FILTER_JOB_PROPERTY_ = 'FORMAT_QA_FILTER_JOB';
var FILTER_TAB_TITLE_ = 'Filtered Issues';
var FILTER_ISSUE_BATCH_SIZE_ = 5;

function filterIssuesByStatus_(status) {
  var html = HtmlService.createHtmlOutput(buildIssueFilterProgressHtml_(status))
    .setWidth(430)
    .setHeight(260);
  DocumentApp.getUi().showModelessDialog(html, 'Filtering issues');
}

function initializeIssueFilterJob(status) {
  var statusKey = normalizeIssueStatus_(status);
  var supportedStatuses = ['open', 'hold', 'fixed', 'not a bug', 'not fixed'];
  if (supportedStatuses.indexOf(statusKey) < 0) {
    throw new Error('Unsupported issue status: ' + status);
  }

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Another issue filter operation is still running. Try again shortly.');
  }
  try {
    var docId = DocumentApp.getActiveDocument().getId();
    var state = readIssueFilterState_(docId);
    var matching = state.issues.filter(function (issue) {
      return normalizeIssueStatus_(issue.status) === statusKey;
    });
    if (!matching.length) {
      return {
        noMatches: true,
        status: status,
        originalUrl: documentTabUrl_(docId, state.tabId)
      };
    }

    var document = DocumentApp.openById(docId);
    var sourceBody = getDocumentTab_(document, state.tabId).getBody();
    var issueEntries = buildIssueBodyEntries_(sourceBody, state.issues);
    var selectedEntries = issueEntries.filter(function (issue) {
      return normalizeIssueStatus_(issue.status) === statusKey;
    });
    var properties = PropertiesService.getUserProperties();
    var filterTabId = getOrCreateFilterTabId_(docId, state, properties);

    var jobId = Utilities.getUuid();
    properties.setProperty(FILTER_JOB_PROPERTY_, JSON.stringify({
      jobId: jobId,
      docId: docId,
      sourceTabId: state.tabId,
      filterTabId: filterTabId,
      status: status,
      total: selectedEntries.length,
      processed: 0,
      prepared: false
    }));
    selectedEntries.forEach(function (issue, index) {
      issue.copyOrder = index;
    });
    return {
      jobId: jobId,
      status: status,
      total: selectedEntries.length,
      processed: 0,
      issues: selectedEntries,
      filterUrl: documentTabUrl_(docId, filterTabId)
    };
  } finally {
    lock.releaseLock();
  }
}

function prepareIssueFilterJob(jobId) {
  var properties = PropertiesService.getUserProperties();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Another issue filter operation is still running. Try again shortly.');
  }
  try {
    var job = readIssueFilterJob_(properties, jobId);
    var document = DocumentApp.openById(job.docId);
    var targetBody = getDocumentTab_(document, job.filterTabId).getBody();
    targetBody.clear();
    var heading = targetBody.appendParagraph(
      'Temporary filtered view: ' + job.status + ' (' + job.total + ' issue(s))'
    );
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    heading.setSpacingAfter(12);
    document.saveAndClose();
    job.prepared = true;
    properties.setProperty(FILTER_JOB_PROPERTY_, JSON.stringify(job));
  } finally {
    lock.releaseLock();
  }
}

function copyIssueFilterBatch(jobId, issueBatch) {
  var properties = PropertiesService.getUserProperties();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Another issue filter operation is still running. Try again shortly.');
  }
  try {
    var job = readIssueFilterJob_(properties, jobId);
    if (!job.prepared) {
      throw new Error('The filtered tab is not ready yet. Start the filter again.');
    }
    var document = DocumentApp.openById(job.docId);
    var sourceBody = getDocumentTab_(document, job.sourceTabId).getBody();
    var targetBody = getDocumentTab_(document, job.filterTabId).getBody();
    if (!Array.isArray(issueBatch) || !issueBatch.length) {
      throw new Error('The issue filter received an empty batch.');
    }
    var batchStart = job.processed;
    issueBatch.forEach(function (issue, batchIndex) {
      if (issue.copyOrder !== batchStart + batchIndex ||
          normalizeIssueStatus_(issue.status) !== normalizeIssueStatus_(job.status)) {
        throw new Error('Issue filter progress is out of sequence. Restart the filter.');
      }
      if (issue.startChildIndex < 0 || issue.endChildIndex > sourceBody.getNumChildren() ||
          issue.headingChildIndex < issue.startChildIndex ||
          issue.headingChildIndex >= issue.endChildIndex) {
        throw new Error('The source issue range changed while filtering. Restart the filter.');
      }
      var sourceHeading = sourceBody.getChild(issue.headingChildIndex);
      var sourceHeadingText = sourceHeading.getType() === DocumentApp.ElementType.LIST_ITEM
        ? sourceHeading.asListItem().getText()
        : sourceHeading.getType() === DocumentApp.ElementType.PARAGRAPH
          ? sourceHeading.asParagraph().getText()
          : '';
      if (!isDocumentIssueHeading_(sourceHeading) ||
          normalizeIssueStatus_(sourceHeadingText) !== normalizeIssueStatus_(issue.title)) {
        throw new Error('The source issue changed while filtering. Restart the filter.');
      }
      for (var i = issue.startChildIndex; i < issue.endChildIndex; i++) {
        if (i === issue.statusChildIndex) continue;
        appendIssueElementCopy_(targetBody, sourceBody.getChild(i));
        if (i === issue.headingChildIndex) {
          appendFilterStatusLabel_(targetBody, issue.status);
        }
      }
      job.processed++;
    });
    document.saveAndClose();
    if (job.processed >= job.total) {
      properties.deleteProperty(FILTER_JOB_PROPERTY_);
    } else {
      properties.setProperty(FILTER_JOB_PROPERTY_, JSON.stringify(job));
    }
    return {
      processed: job.processed,
      total: job.total,
      done: job.processed >= job.total,
      filterUrl: documentTabUrl_(job.docId, job.filterTabId)
    };
  } finally {
    lock.releaseLock();
  }
}

function cancelIssueFilterJob(jobId) {
  var properties = PropertiesService.getUserProperties();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Another issue filter operation is still running. Try again shortly.');
  }
  try {
    var job = readIssueFilterJob_(properties, jobId);
    deleteTemporaryFilterTab_(job.docId, job.filterTabId, job.sourceTabId);
    properties.deleteProperty(FILTER_JOB_PROPERTY_);
    properties.deleteProperty(FILTER_TAB_PROPERTY_);
  } finally {
    lock.releaseLock();
  }
}

function showOriginalIssueDocument_(removeTemporaryTab) {
  var docId = DocumentApp.getActiveDocument().getId();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Another issue filter operation is still running. Try again shortly.');
  }
  var properties = PropertiesService.getUserProperties();
  try {
    var filterTabId = properties.getProperty(FILTER_TAB_PROPERTY_);
    if (removeTemporaryTab && filterTabId) {
      var sourceTabId = readMainDocumentTabId_(docId);
      deleteTemporaryFilterTab_(docId, filterTabId, sourceTabId);
      properties.deleteProperty(FILTER_TAB_PROPERTY_);
      properties.deleteProperty(FILTER_JOB_PROPERTY_);
    }
  } finally {
    lock.releaseLock();
  }
  var mainTabId = readMainDocumentTabId_(docId);
  showIssueNavigationDialog_(
    removeTemporaryTab ? 'Original view restored' : 'All issues',
    removeTemporaryTab
      ? 'The temporary filtered tab was removed. The original issue report is unchanged.'
      : 'The original tab still contains all issues.',
    documentTabUrl_(docId, mainTabId),
    'Open original tab'
  );
}

function readIssueFilterState_(docId) {
  var api = 'https://docs.googleapis.com/v1/documents/' + encodeURIComponent(docId);
  var document = docsApiRequest_(api + '?includeTabsContent=true', 'get', ScriptApp.getOAuthToken());
  var mainTab = document.tabs && document.tabs.length ? document.tabs[0] : null;
  var documentTab = mainTab && mainTab.documentTab;
  var body = documentTab && documentTab.body || document.body || {};
  var content = body.content || [];
  var tabId = mainTab && mainTab.tabProperties && mainTab.tabProperties.tabId;
  var definitions = documentTab && documentTab.dropdownDefinitions || {};
  var definitionsById = {};
  Object.keys(definitions).forEach(function (key) {
    var definition = definitions[key];
    if (definition.dropdownDefinitionId) {
      definitionsById[definition.dropdownDefinitionId] = definition;
    }
  });

  var issues = [];
  for (var i = 0; i < content.length; i++) {
    var item = content[i];
    if (!isApiIssueHeading_(item)) continue;

    var statusDropdown = findDropdownInParagraph_(item.paragraph);
    var statusInFollowingParagraph = false;
    if (!statusDropdown && content[i + 1] && content[i + 1].paragraph &&
        !isApiIssueHeading_(content[i + 1])) {
      statusDropdown = findDropdownInParagraph_(content[i + 1].paragraph);
      statusInFollowingParagraph = !!statusDropdown;
    }
    var status = statusDropdown && getDropdownDisplayValue_(statusDropdown, definitionsById);
    if (!status) {
      throw new Error(
        'Issue ' + (issues.length + 1) + ' has no readable status dropdown. ' +
        'Run Format report > Add issue status dropdowns, then try the filter again.'
      );
    }

    issues.push({
      index: issues.length,
      title: getIssueParagraphText_(item.paragraph),
      status: status,
      statusInFollowingParagraph: statusInFollowingParagraph
    });
  }

  if (!issues.length) {
    throw new Error('No numbered issue headings were found in the document.');
  }
  if (!content.length) {
    throw new Error('The document has no readable content to filter.');
  }

  return {
    tabId: tabId || null,
    tabs: document.tabs || [],
    issues: issues
  };
}

function findDropdownInParagraph_(paragraph) {
  var elements = paragraph.elements || [];
  for (var i = 0; i < elements.length; i++) {
    if (elements[i].dropdown) return elements[i];
  }
  return null;
}

function getDropdownDisplayValue_(element, definitionsById) {
  var props = element.dropdown &&
    element.dropdown.dropdownProperties || {};
  var definition = definitionsById[props.dropdownDefinitionId];
  var options = definition &&
    definition.dropdownDefinitionProperties &&
    definition.dropdownDefinitionProperties.options || [];
  for (var i = 0; i < options.length; i++) {
    if (options[i].optionId === props.selectedOptionId) {
      return options[i].displayValue || null;
    }
  }
  return null;
}

function getIssueParagraphText_(paragraph) {
  return (paragraph.elements || []).map(function (element) {
    return element.textRun ? element.textRun.content : '';
  }).join('').replace(/[\r\n]+$/, '').replace(/\s+/g, ' ').trim();
}

function hasIssueOutlineMarker_(text) {
  return /^\s*(?:\d+[.)]|[A-Za-z][.)]|[IVXLCDM]+[.)]|[ivxlcdm]+[.)]|[•○■])\s+/.test(
    String(text || '')
  );
}

function isApiIssueHeading_(item) {
  if (!item || !item.paragraph) return false;
  if (item.paragraph.bullet) {
    return !(findDropdownInParagraph_(item.paragraph) &&
      !getIssueParagraphText_(item.paragraph));
  }
  var style = item.paragraph.paragraphStyle || {};
  return style.namedStyleType === 'HEADING_2' &&
    hasIssueOutlineMarker_(getIssueParagraphText_(item.paragraph));
}

function isDocumentIssueHeading_(element) {
  if (!element) return false;
  if (element.getType() === DocumentApp.ElementType.LIST_ITEM) return true;
  if (element.getType() !== DocumentApp.ElementType.PARAGRAPH) return false;
  var paragraph = element.asParagraph();
  return paragraph.getHeading() === DocumentApp.ParagraphHeading.HEADING2 &&
    hasIssueOutlineMarker_(paragraph.getText());
}

function normalizeIssueStatus_(status) {
  return String(status || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function readIssueFilterJob_(properties, jobId) {
  var serialized = properties.getProperty(FILTER_JOB_PROPERTY_);
  if (!serialized) throw new Error('The issue filter session has expired. Start the filter again.');
  var job = JSON.parse(serialized);
  if (!job || job.jobId !== jobId) {
    throw new Error('This issue filter is no longer active. Start the filter again.');
  }
  return job;
}

function buildIssueBodyEntries_(body, issues) {
  var headingIndexes = [];
  for (var i = 0; i < body.getNumChildren(); i++) {
    if (isDocumentIssueHeading_(body.getChild(i))) headingIndexes.push(i);
  }
  if (headingIndexes.length !== issues.length) {
    throw new Error('The issue list changed while preparing the filtered view.');
  }

  var starts = headingIndexes.map(function (headingIndex) {
    return headingIndex > 0 &&
      body.getChild(headingIndex - 1).getType() === DocumentApp.ElementType.HORIZONTAL_RULE
      ? headingIndex - 1
      : headingIndex;
  });
  return issues.map(function (issue, index) {
    var headingChildIndex = headingIndexes[index];
    var startChildIndex = starts[index];
    var endChildIndex = index + 1 < starts.length
      ? starts[index + 1]
      : body.getNumChildren();
    var statusChildIndex = issue.statusInFollowingParagraph
      ? headingChildIndex + 1
      : -1;
    if (statusChildIndex >= endChildIndex ||
        (statusChildIndex >= 0 &&
          body.getChild(statusChildIndex).getType() !== DocumentApp.ElementType.PARAGRAPH)) {
      throw new Error('Could not locate the status line for issue ' + (index + 1) + '.');
    }
    return {
      index: issue.index,
      title: issue.title,
      status: issue.status,
      startChildIndex: startChildIndex,
      headingChildIndex: headingChildIndex,
      endChildIndex: endChildIndex,
      statusChildIndex: statusChildIndex
    };
  });
}

function getOrCreateFilterTabId_(docId, state, properties) {
  var existingId = properties.getProperty(FILTER_TAB_PROPERTY_);
  var tabs = flattenDocumentTabs_(state.tabs, []);
  var isExistingTab = tabs.some(function (tab) {
    var props = tab.tabProperties || {};
    return props.tabId === existingId &&
      String(props.title || '').indexOf(FILTER_TAB_TITLE_ + ' ') === 0;
  });
  if (existingId && existingId !== state.tabId && isExistingTab) return existingId;

  var api = 'https://docs.googleapis.com/v1/documents/' +
    encodeURIComponent(docId) + ':batchUpdate';
  var response = docsApiRequest_(api, 'post', ScriptApp.getOAuthToken(), {
    requests: [{
      addDocumentTab: {
        tabProperties: {
          title: FILTER_TAB_TITLE_ + ' ' + Utilities.getUuid().slice(0, 8)
        }
      }
    }]
  });
  var tab = response.replies && response.replies[0] &&
    response.replies[0].addDocumentTab &&
    response.replies[0].addDocumentTab.tabProperties;
  if (!tab || !tab.tabId) {
    throw new Error('Google Docs did not return the temporary filter tab ID.');
  }
  properties.setProperty(FILTER_TAB_PROPERTY_, tab.tabId);
  return tab.tabId;
}

function getDocumentTab_(document, tabId) {
  var tab = null;
  var lookupError = null;
  if (tabId) {
    try {
      tab = document.getTab(tabId);
    } catch (e) {
      lookupError = e;
    }
  } else {
    tab = document.getTabs()[0];
  }
  if (!tab && tabId) {
    tab = findDocumentTabById_(document.getTabs(), tabId);
  }
  if (!tab) {
    var message = 'Could not open the requested document tab' +
      (tabId ? ' (' + tabId + ')' : '') + '.';
    if (lookupError && lookupError.message) message += ' ' + lookupError.message;
    throw new Error(message);
  }
  return tab.asDocumentTab();
}

function findDocumentTabById_(tabs, tabId) {
  for (var i = 0; i < tabs.length; i++) {
    var tab = tabs[i];
    if (tab.getId() === tabId) return tab;
    var childMatch = findDocumentTabById_(tab.getChildTabs(), tabId);
    if (childMatch) return childMatch;
  }
  return null;
}

function appendIssueElementCopy_(body, element) {
  var type = element.getType();
  if (type === DocumentApp.ElementType.PARAGRAPH) {
    body.appendParagraph(element.copy().asParagraph());
  } else if (type === DocumentApp.ElementType.LIST_ITEM) {
    body.appendListItem(element.copy().asListItem());
  } else if (type === DocumentApp.ElementType.TABLE) {
    body.appendTable(element.copy().asTable());
  } else if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
    body.appendHorizontalRule();
  } else if (type === DocumentApp.ElementType.PAGE_BREAK) {
    body.appendPageBreak();
  } else {
    throw new Error('Cannot safely copy issue content of type ' + type + '.');
  }
}

function appendFilterStatusLabel_(body, status) {
  var paragraph = body.appendParagraph(status);
  paragraph.setSpacingBefore(0).setSpacingAfter(5);
  var option = CFG.STATUS_OPTIONS.filter(function (candidate) {
    return normalizeIssueStatus_(candidate.value) === normalizeIssueStatus_(status);
  })[0];
  var attributes = {};
  attributes[DocumentApp.Attribute.FONT_FAMILY] = CFG.FONT;
  attributes[DocumentApp.Attribute.FONT_SIZE] = 10;
  attributes[DocumentApp.Attribute.BOLD] = true;
  attributes[DocumentApp.Attribute.FOREGROUND_COLOR] = option ? option.bg : CFG.ACCENT;
  paragraph.setAttributes(attributes);
}

function deleteTemporaryFilterTab_(docId, filterTabId, sourceTabId) {
  if (!filterTabId || filterTabId === sourceTabId) return;
  var getApi = 'https://docs.googleapis.com/v1/documents/' + encodeURIComponent(docId);
  var document = docsApiRequest_(
    getApi + '?includeTabsContent=true',
    'get',
    ScriptApp.getOAuthToken()
  );
  var temporaryTab = flattenDocumentTabs_(document.tabs || [], []).filter(function (tab) {
    var props = tab.tabProperties || {};
    return props.tabId === filterTabId &&
      String(props.title || '').indexOf(FILTER_TAB_TITLE_ + ' ') === 0;
  })[0];
  if (!temporaryTab) return;
  var api = getApi + ':batchUpdate';
  docsApiRequest_(api, 'post', ScriptApp.getOAuthToken(), {
    requests: [{ deleteTab: { tabId: filterTabId } }]
  });
}

function readMainDocumentTabId_(docId) {
  var api = 'https://docs.googleapis.com/v1/documents/' + encodeURIComponent(docId);
  var document = docsApiRequest_(
    api + '?includeTabsContent=true',
    'get',
    ScriptApp.getOAuthToken()
  );
  var tab = document.tabs && document.tabs[0];
  return tab && tab.tabProperties && tab.tabProperties.tabId || null;
}

function documentTabUrl_(docId, tabId) {
  var url = 'https://docs.google.com/document/d/' +
    encodeURIComponent(docId) + '/edit';
  return tabId ? url + '?tab=' + encodeURIComponent(tabId) : url;
}

function showIssueNavigationDialog_(title, message, url, linkLabel) {
  var html = '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>body{font:13px Arial,sans-serif;color:#202124;padding:18px}p{line-height:1.5}' +
    'a{color:#1a73e8;font-weight:bold}</style></head><body>' +
    '<p>' + escapeHtml_(message) + '</p><p><a href="' + url + '">' +
    escapeHtml_(linkLabel) + '</a></p>' +
    '<p><button onclick="google.script.host.close()">Close</button></p></body></html>';
  DocumentApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(390).setHeight(180),
    title
  );
}

function buildIssueFilterProgressHtml_(status) {
  return '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font:13px Arial,sans-serif;color:#202124;padding:18px}' +
    '#track{height:14px;background:#e8eaed;border-radius:8px;overflow:hidden;margin:14px 0}' +
    '#bar{height:100%;width:0;background:#1a73e8;transition:width .2s}' +
    '#link{display:none;margin-top:16px;font-weight:bold;color:#1a73e8}' +
    '#error{color:#b3261e;white-space:pre-wrap}button{margin-top:16px}</style></head><body>' +
    '<div id="message">Preparing the filtered view…</div>' +
    '<div id="track"><div id="bar"></div></div><div id="percent">0%</div>' +
    '<a id="link"></a><div id="error"></div><button id="close" disabled onclick="closeDialog()">Close</button>' +
    '<script>var status=' + JSON.stringify(status) + ',job=null,finished=false;' +
    'function start(){google.script.run.withSuccessHandler(begin).withFailureHandler(fail).initializeIssueFilterJob(status);}' +
    'function begin(result){document.getElementById("close").disabled=false;if(result.noMatches){finished=true;document.getElementById("message").textContent="No issues have status "+status+".";showLink(result.originalUrl,"Open original tab");return;}' +
    'job=result;document.getElementById("message").textContent="Preparing the filtered tab…";' +
    'google.script.run.withSuccessHandler(function(){document.getElementById("message").textContent="Copying "+status+" issues…";runBatch();})' +
    '.withFailureHandler(fail).prepareIssueFilterJob(job.jobId);}' +
    'function runBatch(){var batch=job.issues.slice(job.processed,job.processed+' +
      FILTER_ISSUE_BATCH_SIZE_ + ');google.script.run.withSuccessHandler(update).withFailureHandler(fail).copyIssueFilterBatch(job.jobId,batch);}' +
    'function update(result){job.processed=result.processed;var percent=Math.floor(100*result.processed/result.total);' +
    'document.getElementById("bar").style.width=percent+"%";document.getElementById("percent").textContent=percent+"%";' +
    'document.getElementById("message").textContent="Copied "+result.processed+" of "+result.total+" issues";' +
    'if(result.done){finished=true;document.getElementById("message").textContent="Filtered view ready: "+result.total+" issue(s) with status "+status+".";showLink(result.filterUrl,"Open filtered view in this document");}' +
    'else{setTimeout(runBatch,80);}}' +
    'function showLink(url,label){var a=document.getElementById("link");a.href=url;a.textContent=label;a.style.display="block";}' +
    'function fail(error){finished=true;document.getElementById("close").disabled=false;document.getElementById("message").textContent="Filtering stopped.";document.getElementById("error").textContent=error&&error.message?error.message:String(error);' +
    'if(job&&job.jobId)google.script.run.withSuccessHandler(function(){}).cancelIssueFilterJob(job.jobId);}' +
    'function closeDialog(){if(job&&!finished)google.script.run.withSuccessHandler(function(){google.script.host.close();}).cancelIssueFilterJob(job.jobId);else google.script.host.close();}' +
    'start();</script></body></html>';
}

function flattenDocumentTabs_(tabs, result) {
  (tabs || []).forEach(function (tab) {
    result.push(tab);
    flattenDocumentTabs_(tab.childTabs, result);
  });
  return result;
}

function escapeHtml_(value) {
  return String(value).replace(/[&<>"']/g, function (character) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[character];
  });
}

function buildDialogHtml_() {
  var bgPresets = ['#F4F6FA', '#FFF8E7', '#EEF7F0', '#FDF0F0', '#F3EFFA', '#FFFFFF'];
  var accPresets = ['#1F3A5F', '#0F766E', '#7C2D12', '#4C1D95', '#9F1239', '#111827'];
  function sw(list, target) {
    return list.map(function (c) {
      return '<span class="sw" style="background:' + c + '" title="' + c + '" onclick="pick(\'' + target + '\',\'' + c + '\')"></span>';
    }).join('');
  }
  return '<!DOCTYPE html><html><head><base target="_top"><style>' +
  'body{font-family:Arial,sans-serif;font-size:13px;color:#1F2933;margin:0;padding:18px 20px}' +
  'h3{margin:0 0 4px;font-size:15px}.hint{color:#5F6368;margin:0 0 14px}' +
  'fieldset{border:1px solid #D0D7E2;border-radius:8px;margin:0 0 12px;padding:10px 12px}' +
  'legend{font-weight:bold;padding:0 6px;color:#1F3A5F}label{display:block;margin:5px 0}' +
  'select{width:100%;padding:6px;font-size:13px;margin-top:4px}' +
  '.row{display:flex;align-items:center;gap:10px;margin-top:6px}' +
  'input[type=color]{width:46px;height:34px;padding:0;border:1px solid #D0D7E2;border-radius:6px;background:none;cursor:pointer}' +
  '.sw{display:inline-block;width:26px;height:26px;border-radius:50%;border:1px solid #9AA5B1;margin-right:6px;cursor:pointer}' +
  '.prev{flex:1;height:34px;border-radius:6px;border:1px solid #D0D7E2}' +
  'button{padding:9px 18px;font-size:13px;border-radius:6px;border:0;cursor:pointer}' +
  '#go{background:#1F3A5F;color:#fff}#cancel{background:#E8EAED;margin-left:8px}' +
  '#msg{margin-top:10px;color:#B3261E}#bgbox.off{opacity:.4;pointer-events:none}' +
  '</style></head><body>' +
  '<h3>Format QA report</h3><p class="hint">Choose how the document should look, then click Format document.</p>' +

  '<fieldset><legend>List style</legend>' +
  '<select id="listType">' +
  '<option value="NUMBER">1. 2. 3. – Numbers</option>' +
  '<option value="BULLET">● Bullets</option>' +
  '<option value="HOLLOW_BULLET">○ Hollow bullets</option>' +
  '<option value="SQUARE_BULLET">■ Square bullets</option>' +
  '<option value="LATIN_UPPER">A. B. C. – Capital letters</option>' +
  '<option value="LATIN_LOWER">a. b. c. – Small letters</option>' +
  '<option value="ROMAN_UPPER">I. II. III. – Roman capitals</option>' +
  '<option value="ROMAN_LOWER">i. ii. iii. – Roman small</option>' +
  '</select></fieldset>' +

  '<fieldset><legend>Document outline</legend>' +
  '<label><input type="radio" name="outline" value="1" checked> Yes, add issue headings to the outline</label>' +
  '<label><input type="radio" name="outline" value="0"> No, leave the current heading styles unchanged</label></fieldset>' +

  '<fieldset><legend>Separator line between issues</legend>' +
  '<label><input type="radio" name="div" value="1" checked> Yes, add a thin line</label>' +
  '<label><input type="radio" name="div" value="0"> No lines</label></fieldset>' +

  '<fieldset><legend>Image size</legend>' +
  '<label><input type="radio" name="img" value="0" checked> Keep images exactly as they are</label>' +
  '<label><input type="radio" name="img" value="1"> Shrink only images wider than the page</label></fieldset>' +

  '<fieldset><legend>Page background</legend>' +
  '<label><input type="radio" name="bg" value="1" checked onchange="toggleBg()"> Add background colour</label>' +
  '<label><input type="radio" name="bg" value="0" onchange="toggleBg()"> Keep white</label>' +
  '<div id="bgbox"><div class="row"><input type="color" id="bgColor" value="' + CFG.PAGE_BG + '" oninput="prev()">' +
  '<div class="prev" id="bgPrev"></div></div><div class="row">' + sw(bgPresets, 'bgColor') + '</div></div></fieldset>' +

  '<fieldset><legend>Banner colour</legend>' +
  '<div class="row"><input type="color" id="accent" value="' + CFG.ACCENT + '" oninput="prev()">' +
  '<div class="prev" id="accPrev"></div></div><div class="row">' + sw(accPresets, 'accent') + '</div></fieldset>' +

  '<button id="go" onclick="go()">Format document</button><button id="cancel" onclick="google.script.host.close()">Cancel</button>' +
  '<div id="msg"></div>' +
  '<script>' +
  'function pick(id,c){document.getElementById(id).value=c;prev();}' +
  'function prev(){document.getElementById("bgPrev").style.background=document.getElementById("bgColor").value;' +
  'document.getElementById("accPrev").style.background=document.getElementById("accent").value;}' +
  'function val(n){return document.querySelector("input[name="+n+"]:checked").value;}' +
  'function toggleBg(){document.getElementById("bgbox").className=(val("bg")==="1")?"":"off";}' +
  'function go(){var b=document.getElementById("go");b.disabled=true;b.textContent="Formatting… please wait";' +
  'var o={listType:document.getElementById("listType").value,dividers:val("div")==="1",fitImages:val("img")==="1",pageBg:val("bg")==="1",' +
  'bgColor:document.getElementById("bgColor").value,accent:document.getElementById("accent").value,makeOutline:val("outline")==="1"};' +
  'google.script.run.withSuccessHandler(function(r){if(r==="done"||r==="opened-dialog"){google.script.host.close();}else{b.disabled=false;b.textContent="Format document";document.getElementById("msg").textContent=r;}})' +
  '.withFailureHandler(function(e){b.disabled=false;b.textContent="Format document";document.getElementById("msg").textContent="Error: "+e.message;})' +
  '.applyFormatting(o);}' +
  'prev();' +
  '</script></body></html>';
}

/* ---------- page + defaults ---------- */
function setPageAndBaseStyle_(body) {
  body.setPageWidth(CFG.PAGE_W).setPageHeight(CFG.PAGE_H);
  body.setMarginTop(CFG.MARGIN_PT + 10).setMarginBottom(CFG.MARGIN_PT + 10)
      .setMarginLeft(CFG.MARGIN_PT).setMarginRight(CFG.MARGIN_PT);
  var base = {};
  base[DocumentApp.Attribute.FONT_FAMILY] = CFG.FONT;
  base[DocumentApp.Attribute.FONT_SIZE] = CFG.BODY_PT;
  base[DocumentApp.Attribute.FOREGROUND_COLOR] = CFG.TEXT_COLOR;
  base[DocumentApp.Attribute.LINE_SPACING] = 1.3;
  body.setAttributes(base);
}

function setPageBackground_(docId, color) {
  // Page colour is only available through the Docs API. Works with OR without the
  // "Google Docs API" service: uses the service if enabled, otherwise calls the REST API.
  var req = { requests: [{
    updateDocumentStyle: {
      documentStyle: { background: { color: { color: { rgbColor: hexToRgb_(color || CFG.PAGE_BG) } } } },
      fields: 'background'
    } }] };
  try {
    if (typeof Docs !== 'undefined') {
      Docs.Documents.batchUpdate(req, docId);
    } else {
      var res = UrlFetchApp.fetch('https://docs.googleapis.com/v1/documents/' + docId + ':batchUpdate', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: JSON.stringify(req),
        muteHttpExceptions: true
      });
      if (res.getResponseCode() !== 200) Logger.log('Page background failed: ' + res.getContentText());
    }
  } catch (e) {
    Logger.log('Page background skipped: ' + e);
  }
}

function hexToRgb_(hex) {
  var h = hex.replace('#', '');
  return { red: parseInt(h.substr(0, 2), 16) / 255,
           green: parseInt(h.substr(2, 2), 16) / 255,
           blue: parseInt(h.substr(4, 2), 16) / 255 };
}

/** Canonical key for a dropdown option's textStyle colours, for change detection. */
function styleKey_(textStyle) {
  if (!textStyle) return '';
  function k(field) {
    var c = textStyle[field] && textStyle[field].color && textStyle[field].color.rgbColor;
    return c ? [c.red, c.green, c.blue].map(function (v) { return Math.round(v * 100) / 100; }).join(',') : '-';
  }
  return k('foregroundColor') + '|' + k('backgroundColor');
}

/* ---------- banner (replaces the plain title) ---------- */
function isBanner_(tbl) {
  return tbl.getNumRows() === 1 && tbl.getRow(0).getNumCells() === 1 &&
         tbl.getCell(0, 0).getText().indexOf('issues found') > -1;
}

function buildBanner_(body, issueCount) {
  var titleText = CFG.HEADER_TEXT;
  var first = body.getChild(0);
  // a banner from an earlier run: reuse its title, then remove it so it is never duplicated
  if (first.getType() === DocumentApp.ElementType.TABLE && isBanner_(first.asTable())) {
    titleText = first.asTable().getCell(0, 0).getChild(0).asParagraph().getText() || titleText;
    body.removeChild(first);
    first = body.getChild(0);
  }
  var hadTitle = first.getType() === DocumentApp.ElementType.PARAGRAPH &&
                 first.asParagraph().getText().replace(/\s/g, '') !== '';
  if (hadTitle) titleText = first.asParagraph().getText().replace(/[\s\u000b]+$/, '');

  var tbl = body.insertTable(0, [[titleText]]);
  tbl.setBorderWidth(0);
  tbl.setColumnWidth(0, CFG.PAGE_W - 2 * CFG.MARGIN_PT);
  var cell = tbl.getCell(0, 0);
  cell.setBackgroundColor(CFG.BANNER_BG)
      .setPaddingTop(22).setPaddingBottom(22).setPaddingLeft(24).setPaddingRight(24);

  var tp = cell.getChild(0).asParagraph();
  tp.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  tp.setIndentStart(0).setIndentFirstLine(0).setSpacingBefore(0).setSpacingAfter(6);
  tp.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  var ts = {};
  ts[DocumentApp.Attribute.FONT_FAMILY] = CFG.FONT;
  ts[DocumentApp.Attribute.FONT_SIZE] = CFG.BANNER_TITLE_PT;
  ts[DocumentApp.Attribute.BOLD] = true;
  ts[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FFFFFF';
  tp.setAttributes(ts);

  var sub = cell.appendParagraph(issueCount + ' issues found  •  Reported by QA');
  sub.setIndentStart(0).setIndentFirstLine(0).setSpacingBefore(0).setSpacingAfter(0);
  var ss = {};
  ss[DocumentApp.Attribute.FONT_FAMILY] = CFG.FONT;
  ss[DocumentApp.Attribute.FONT_SIZE] = 11;
  ss[DocumentApp.Attribute.BOLD] = false;
  ss[DocumentApp.Attribute.FOREGROUND_COLOR] = CFG.BANNER_SUB_COLOR;
  sub.setAttributes(ss);

  // remove the original title paragraph (now just after the table) and add a small spacer
  if (hadTitle) body.removeChild(body.getChild(1));
  var spacer = body.insertParagraph(1, '');
  spacer.setSpacingBefore(0).setSpacingAfter(0);
  spacer.editAsText().setFontSize(4);
}

/* ---------- cleanup ---------- */
function countIssues_(body) {
  var n = 0;
  for (var i = 0; i < body.getNumChildren(); i++) {
    if (isDocumentIssueHeading_(body.getChild(i))) n++;
  }
  return n;
}

function tidyEmptyParagraphs_(body) {
  for (var i = body.getNumChildren() - 1; i >= 0; i--) {
    var el = body.getChild(i);
    var type = el.getType();
    if (type === DocumentApp.ElementType.LIST_ITEM) { stripTrailingBreaks_(el.asListItem()); continue; }
    if (type !== DocumentApp.ElementType.PARAGRAPH) continue;
    var p = el.asParagraph();
    stripTrailingBreaks_(p);
    if (!containsImage_(p) && p.getText().replace(/\s/g, '') === '' && body.getNumChildren() > 1) {
      try { body.removeChild(el); } catch (e) { /* last paragraph can't be removed */ }
    }
  }
}

function containsImage_(p) {
  for (var i = 0; i < p.getNumChildren(); i++) {
    if (p.getChild(i).getType() === DocumentApp.ElementType.INLINE_IMAGE) return true;
  }
  return false;
}

function stripTrailingBreaks_(p) {
  for (var i = p.getNumChildren() - 1; i >= 0; i--) {
    var c = p.getChild(i), t = c.getType();
    if (t === DocumentApp.ElementType.TEXT) {
      var txt = c.asText(), s = txt.getText();
      var trimmed = s.replace(/[\s\u000b\r\n]+$/, '');
      if (trimmed.length === 0) { if (p.getNumChildren() > 1) c.removeFromParent(); else break; }
      else { if (trimmed.length < s.length) txt.deleteText(trimmed.length, s.length - 1); break; }
    } else if (t === DocumentApp.ElementType.INLINE_IMAGE) {
      break;
    } else {
      c.removeFromParent();
    }
  }
}

/* ---------- numbered issues ---------- */
function outlineMarkerForIssue_(number, listType) {
  if (listType === 'BULLET') return '• ';
  if (listType === 'HOLLOW_BULLET') return '○ ';
  if (listType === 'SQUARE_BULLET') return '■ ';
  if (listType === 'LATIN_UPPER' || listType === 'LATIN_LOWER') {
    var letter = '';
    var value = number;
    while (value > 0) {
      value--;
      letter = String.fromCharCode(65 + value % 26) + letter;
      value = Math.floor(value / 26);
    }
    return (listType === 'LATIN_LOWER' ? letter.toLowerCase() : letter) + '. ';
  }
  if (listType === 'ROMAN_UPPER' || listType === 'ROMAN_LOWER') {
    var romanValues = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
    var romanSymbols = ['M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I'];
    var roman = '';
    var remaining = number;
    for (var i = 0; i < romanValues.length; i++) {
      while (remaining >= romanValues[i]) {
        roman += romanSymbols[i];
        remaining -= romanValues[i];
      }
    }
    if (listType === 'ROMAN_LOWER') roman = roman.toLowerCase();
    return roman + '. ';
  }
  return number + '. ';
}

function createIssueOutline_(body, listType) {
  if (body.getNumChildren() &&
      body.getChild(0).getType() === DocumentApp.ElementType.TABLE &&
      isBanner_(body.getChild(0).asTable())) {
    var title = body.getChild(0).asTable().getCell(0, 0).getChild(0).asParagraph();
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    title.setIndentStart(0).setIndentFirstLine(0).setSpacingBefore(0).setSpacingAfter(6);
    title.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    var titleStyle = {};
    titleStyle[DocumentApp.Attribute.FONT_FAMILY] = CFG.FONT;
    titleStyle[DocumentApp.Attribute.FONT_SIZE] = CFG.BANNER_TITLE_PT;
    titleStyle[DocumentApp.Attribute.BOLD] = true;
    titleStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FFFFFF';
    title.setAttributes(titleStyle);
  }

  var issueCount = 0;
  for (var i = 0; i < body.getNumChildren(); i++) {
    var element = body.getChild(i);
    if (!isDocumentIssueHeading_(element)) continue;
    var isListItem = element.getType() === DocumentApp.ElementType.LIST_ITEM;
    var item = isListItem ? element.asListItem() : element.asParagraph();
    var text = item.editAsText();
    if (!hasIssueOutlineMarker_(text.getText())) {
      text.insertText(0, outlineMarkerForIssue_(issueCount + 1, listType));
    }
    item.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    if (isListItem) item.setGlyphType(CFG.GLYPH || DocumentApp.GlyphType.NUMBER);
    item.setIndentStart(24).setIndentFirstLine(12);
    item.setSpacingBefore(CFG.ISSUE_SPACE_BEFORE).setSpacingAfter(CFG.ISSUE_SPACE_AFTER);
    item.setLineSpacing(1.35);
    item.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    text = item.editAsText();
    var textLength = text.getText().length;
    if (textLength) {
      text.setFontFamily(0, textLength - 1, CFG.FONT);
      text.setFontSize(0, textLength - 1, CFG.BODY_PT);
      text.setForegroundColor(0, textLength - 1, CFG.TEXT_COLOR);
      text.setBold(0, textLength - 1, false);
      boldIssueSentence_(item);
    }
    issueCount++;
  }
  if (!issueCount) throw new Error('No numbered issue headings were found.');
  return 'Created outline entries for ' + issueCount + ' issue(s).';
}

function applyIssueOutlineStyles_(docId) {
  var api = 'https://docs.googleapis.com/v1/documents/' + encodeURIComponent(docId);
  var document = docsApiRequest_(
    api + '?includeTabsContent=true',
    'get',
    ScriptApp.getOAuthToken()
  );
  var mainTab = document.tabs && document.tabs[0];
  var content = mainTab && mainTab.documentTab && mainTab.documentTab.body
    ? mainTab.documentTab.body.content || []
    : document.body && document.body.content || [];
  var tabId = mainTab && mainTab.tabProperties && mainTab.tabProperties.tabId;
  var headingGroups = [];
  content.forEach(function (item) {
    if (!item.paragraph ||
        !hasIssueOutlineMarker_(getIssueParagraphText_(item.paragraph))) return;
    var isListItem = !!item.paragraph.bullet;
    var style = item.paragraph.paragraphStyle || {};
    if (!isListItem && style.namedStyleType !== 'HEADING_2') return;

    var range = { startIndex: item.startIndex, endIndex: item.endIndex };
    if (tabId) range.tabId = tabId;
    var requests = [];
    if (isListItem) requests.push({ deleteParagraphBullets: { range: range } });
    requests.push({
      updateParagraphStyle: {
        range: range,
        paragraphStyle: { namedStyleType: 'HEADING_2' },
        fields: 'namedStyleType'
      }
    });
    headingGroups.push({ startIndex: item.startIndex, requests: requests });
  });
  headingGroups.sort(function (a, b) {
    return b.startIndex - a.startIndex;
  });
  var requests = [];
  headingGroups.forEach(function (group) {
    Array.prototype.push.apply(requests, group.requests);
  });
  for (var i = 0; i < requests.length; i += 200) {
    docsApiRequest_(api + ':batchUpdate', 'post', ScriptApp.getOAuthToken(), {
      requests: requests.slice(i, i + 200)
    });
  }
  return 'Applied heading styles to ' + headingGroups.length +
    ' numbered issue outline entry(ies).';
}

function styleIssues_(body) {
  for (var i = 0; i < body.getNumChildren(); i++) {
    var el = body.getChild(i), type = el.getType();
    if (isDocumentIssueHeading_(el)) {
      var item = type === DocumentApp.ElementType.LIST_ITEM
        ? el.asListItem()
        : el.asParagraph();
      if (type === DocumentApp.ElementType.LIST_ITEM) {
        item.setGlyphType(CFG.GLYPH || DocumentApp.GlyphType.NUMBER);
        cleanLeadingSpaces_(item);
      }
      item.setIndentStart(24).setIndentFirstLine(12);
      item.setSpacingBefore(CFG.ISSUE_SPACE_BEFORE).setSpacingAfter(CFG.ISSUE_SPACE_AFTER);
      item.setLineSpacing(1.35);
      item.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      boldIssueSentence_(item);
      resizeImages_(item, CFG.IMG_MAX_W);
    } else if (type === DocumentApp.ElementType.PARAGRAPH) {
      var p = el.asParagraph();
      if (containsImage_(p)) resizeImages_(p, CFG.IMG_MAX_W);
    }
  }
}

// thin line before every issue except the first, so each bullet + its screenshot reads as one block
function addIssueDividers_(body) {
  var seenFirst = false;
  for (var i = 0; i < body.getNumChildren(); i++) {
    if (!isDocumentIssueHeading_(body.getChild(i))) continue;
    if (!seenFirst) { seenFirst = true; continue; }
    body.insertHorizontalRule(i);
    i++;                                   // skip over the rule we just added
  }
}

function cleanLeadingSpaces_(p) {
  if (p.getNumChildren() === 0) return;
  var c = p.getChild(0);
  if (c.getType() === DocumentApp.ElementType.TEXT) {
    var t = c.asText(), m = t.getText().match(/^\s+/);
    if (m && m[0].length < t.getText().length) t.deleteText(0, m[0].length - 1);
  }
}

// first sentence/line of each issue = bold headline in the accent colour
function boldIssueSentence_(p) {
  if (p.getNumChildren() === 0) return;
  var c = p.getChild(0);
  if (c.getType() !== DocumentApp.ElementType.TEXT) return;
  var t = c.asText(), s = t.getText();
  if (!s.length) return;
  var end = s.search(/[\u000b\r\n]/);
  if (end === -1) end = s.length;
  if (end > 0) {
    t.setBold(0, end - 1, true);
    t.setForegroundColor(0, end - 1, CFG.HEADLINE_COLOR);
  }
}

/* ---------- images ---------- */
function resizeImages_(container, maxW) {
  if (!CFG.FIT_IMAGES) return;     // user chose to keep images exactly as they are
  // Keeps the image exactly as it is unless it is wider than the space available.
  for (var i = 0; i < container.getNumChildren(); i++) {
    var c = container.getChild(i);
    if (c.getType() !== DocumentApp.ElementType.INLINE_IMAGE) continue;
    var img = c.asInlineImage(), w = img.getWidth(), h = img.getHeight();
    if (!w || !h || w <= maxW) continue;
    var scale = maxW / w;
    img.setWidth(Math.round(w * scale)).setHeight(Math.round(h * scale));
  }
}

/* ---------- links ---------- */
function linkifyUrls_(body) {
  var re = /https?:\/\/[^\s)\]>"']+/g;
  var text = body.editAsText(), s = text.getText(), m;
  while ((m = re.exec(s)) !== null) {
    var start = m.index, end = start + m[0].length - 1;
    while (end > start && /[.,;:!?]/.test(s.charAt(end))) end--;
    text.setLinkUrl(start, end, s.substring(start, end + 1));
    text.setForegroundColor(start, end, CFG.LINK_COLOR);
    text.setUnderline(start, end, true);
    text.setBold(start, end, false);
  }
}

/* ---------- screenshot tables -> white cards with an accent bar ---------- */
function styleScreenshotTables_(body) {
  var tables = body.getTables();
  for (var t = 0; t < tables.length; t++) {
    var tbl = tables[t];
    if (tbl.getNumRows() > 1 && rowIsEmpty_(tbl.getRow(0))) tbl.removeRow(0);   // old accent bar from a previous run
    tbl.setBorderWidth(1);
    tbl.setBorderColor(CFG.CARD_BORDER);
    var cols = tbl.getRow(0).getNumCells();
    var cellW = Math.floor((CFG.PAGE_W - 2 * CFG.MARGIN_PT) / cols);
    var pad = 10;
    for (var r = 0; r < tbl.getNumRows(); r++) {
      var row = tbl.getRow(r);
      for (var c = 0; c < row.getNumCells(); c++) {
        var cell = row.getCell(c);
        cell.setWidth(cellW);
        cell.setBackgroundColor(CFG.CARD_BG);
        cell.setPaddingTop(pad).setPaddingBottom(pad).setPaddingLeft(pad).setPaddingRight(pad);
        cell.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);
        for (var k = 0; k < cell.getNumChildren(); k++) {
          var ch = cell.getChild(k);
          if (ch.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
          var p = ch.asParagraph();
          p.setIndentStart(0).setIndentFirstLine(0).setSpacingBefore(0).setSpacingAfter(0);
          p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          stripTrailingBreaks_(p);
          resizeImages_(p, cellW - 2 * pad - 4);
        }
      }
    }
    addAccentBar_(tbl, cols);
    spaceAround_(body, tbl);
  }
}

// thin coloured strip across the top of the table (clone of row 0, emptied, so columns always match)
function rowIsEmpty_(row) {
  for (var c = 0; c < row.getNumCells(); c++) {
    var cell = row.getCell(c);
    if (cell.getText().replace(/\s/g, '') !== '') return false;
    for (var k = 0; k < cell.getNumChildren(); k++) {
      var ch = cell.getChild(k);
      if (ch.getType() === DocumentApp.ElementType.PARAGRAPH && containsImage_(ch.asParagraph())) return false;
    }
  }
  return true;
}

function addAccentBar_(tbl, cols) {
  var bar = tbl.insertTableRow(0, tbl.getRow(0).copy());
  for (var i = 0; i < bar.getNumCells(); i++) {
    var cell = bar.getCell(i);
    cell.clear();
    cell.setBackgroundColor(CFG.ACCENT);
    cell.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
    var p = cell.getChild(0).asParagraph();
    p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    p.editAsText().setFontSize(2);
  }
  bar.setMinimumHeight(4);
}

// small gap above the table so it doesn't touch the text before it
function spaceAround_(body, tbl) {
  var idx = body.getChildIndex(tbl);
  var sp = body.insertParagraph(idx, '');
  sp.setSpacingBefore(0).setSpacingAfter(0);
  sp.editAsText().setFontSize(6);
}

/* ---------- header / footer ---------- */
function addHeaderFooter_(doc) {
  var header = doc.getHeader() || doc.addHeader();
  header.clear();
  var hp = header.getParagraphs()[0] || header.appendParagraph('');
  hp.setText(CFG.HEADER_TEXT);
  hp.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  hp.setAttributes(small_(9, CFG.ACCENT, true));

  var footer = doc.getFooter() || doc.addFooter();
  footer.clear();
  var fp = footer.getParagraphs()[0] || footer.appendParagraph('');
  fp.setText(CFG.FOOTER_TEXT);
  fp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  fp.setAttributes(small_(9, '#5F6368', false));
}

function small_(size, color, bold) {
  var a = {};
  a[DocumentApp.Attribute.FONT_FAMILY] = CFG.FONT;
  a[DocumentApp.Attribute.FONT_SIZE] = size;
  a[DocumentApp.Attribute.FOREGROUND_COLOR] = color;
  a[DocumentApp.Attribute.BOLD] = bold;
  return a;
}