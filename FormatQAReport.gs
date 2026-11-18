/**
 * Format QA Issue Report v2  –  Google Apps Script (bound to a Google Doc)
 *
 * WHAT IT DOES
 *  - Dark-blue title banner (white text) with issue count + date
 *  - Soft page background colour
 *  - Screenshots scaled to fit, side-by-side screenshot tables turned into white "cards"
 *  - Clean numbered issues: bold headline sentence, consistent spacing, clickable links
 *  - Native Google Docs status dropdown chips beside each numbered issue
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
  DATE_LABEL: '4–5 Nov',
  FOOTER_TEXT: 'Internal – QA findings'   // page numbers: Insert > Page numbers
};

/** Adds a "Format report" menu to the Doc (appears after you reload it). */
function onOpen() {
  DocumentApp.getUi().createMenu('Format report')
    .addItem('Format QA report…', 'formatQAReport')
    .addItem('Add issue status dropdowns', 'addIssueStatusDropdowns')
    .addToUi();
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
  var r = applyFormatting({ listType: 'NUMBER', dividers: true, fitImages: false, pageBg: true, bgColor: CFG.PAGE_BG, accent: CFG.ACCENT });
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
    catch (e) { failed.push(name); Logger.log('FAILED  ' + name + ' -> ' + e); }
  }

  var issueCount = countIssues_(body);
  step('page + base style',   function () { setPageAndBaseStyle_(body); });
  step('tidy blank lines',    function () { tidyEmptyParagraphs_(body); });
  step('style issues/images', function () { styleIssues_(body); });
  if (opts.dividers) step('separator lines', function () { addIssueDividers_(body); });
  step('links',               function () { linkifyUrls_(body); });
  step('screenshot tables',   function () { styleScreenshotTables_(body); });
  step('title banner',        function () { buildBanner_(body, issueCount); });
  step('header + footer',     function () { addHeaderFooter_(doc); });

  doc.saveAndClose();
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
    if (doc.tabs && doc.tabs.length && doc.tabs[0].documentTab) {
      var tab = doc.tabs[0].documentTab;
      if (tab.body) content = tab.body.content || content;
      definitionMap = tab.dropdownDefinitions || {};
    }
    return { content: content, definitions: definitionMap };
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
  state.content.forEach(function (item) {
    if (!item.paragraph || !item.paragraph.bullet) return;
    issueElements.push(item);

    var dropdown = (item.paragraph.elements || []).filter(function (element) {
      return !!element.dropdown;
    })[0];
    var status = options[0].displayValue;
    if (dropdown) {
      var props = dropdown.dropdown.dropdownProperties || {};
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
    if (!item.paragraph || !item.paragraph.bullet) return;
    var elements = item.paragraph.elements || [];
    var image = null;
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].inlineObjectElement) {
        image = elements[i];
        break;
      }
    }

    var index = image ? image.startIndex : item.endIndex - 1;
    var needsSpace = true;
    for (var j = 0; j < elements.length; j++) {
      var element = elements[j];
      if (!element.textRun || element.endIndex > index) continue;
      var precedingText = element.textRun.content.replace(/[\r\n]+$/, '');
      if (/\s$/.test(precedingText)) needsSpace = false;
    }
    insertionPlans.push({
      index: index,
      needsSpace: needsSpace,
      statusOptionId: desiredByName[issueStatuses[statusIndex]]
    });
    statusIndex++;
  });

  insertionPlans.sort(function (a, b) { return b.index - a.index; });
  var insertRequests = [];
  insertionPlans.forEach(function (plan) {
    var index = plan.index;
    if (plan.needsSpace) {
      insertRequests.push({
        insertText: { location: { index: index }, text: ' ' }
      });
      index++;
    }
    insertRequests.push({
      insertDropdown: {
        location: { index: index },
        dropdownDefinitionId: dropdownDefinitionId,
        selectedOptionId: plan.statusOptionId
      }
    });
  });

  docsApiRequest_(api + ':batchUpdate', 'post', token, {
    requests: insertRequests
  });
  return 'Positioned status dropdowns after the issue text for ' +
    issueElements.length + ' issue(s).';
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
  'bgColor:document.getElementById("bgColor").value,accent:document.getElementById("accent").value};' +
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

  var sub = cell.appendParagraph(issueCount + ' issues found  •  ' + CFG.DATE_LABEL + '  •  Reported by QA');
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
    if (body.getChild(i).getType() === DocumentApp.ElementType.LIST_ITEM) n++;
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
function styleIssues_(body) {
  for (var i = 0; i < body.getNumChildren(); i++) {
    var el = body.getChild(i), type = el.getType();
    if (type === DocumentApp.ElementType.LIST_ITEM) {
      var li = el.asListItem();
      li.setGlyphType(CFG.GLYPH || DocumentApp.GlyphType.NUMBER);
      li.setIndentStart(24).setIndentFirstLine(12);
      li.setSpacingBefore(CFG.ISSUE_SPACE_BEFORE).setSpacingAfter(CFG.ISSUE_SPACE_AFTER);
      li.setLineSpacing(1.35);
      li.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      cleanLeadingSpaces_(li);
      boldIssueSentence_(li);
      resizeImages_(li, CFG.IMG_MAX_W);
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
    if (body.getChild(i).getType() !== DocumentApp.ElementType.LIST_ITEM) continue;
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
  hp.setText(CFG.HEADER_TEXT + '   |   ' + CFG.DATE_LABEL);
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