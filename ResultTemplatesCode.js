// ============================================================
// ResultTemplatesCode.gs
// X-Ray and Lab result generation using templates
// ============================================================

var XRAY_TEMPLATE_FOLDER = 'A-Lab X-Ray Templates';
var XRAY_TEMPLATE_FILENAME = 'xray_result_template.docx';
var XRAY_RESULTS_FOLDER = 'A-Lab Results';

// ── TEMPLATE CONFIG (sections) ──────────────────────────────
function getXrayTemplateConfig() {
  try {
    var raw = getSettingValue_('xray_template_config', '');
    if (raw) return JSON.parse(raw);
  } catch (e) { }
  return {
    sections: [
      { id: 'clinical_data', label: 'CLINICAL DATA', required: false },
      { id: 'findings', label: 'FINDINGS', required: true },
      { id: 'impression', label: 'IMPRESSION', required: true }
    ]
  };
}

function saveXrayTemplateConfig(config) {
  try {
    var user = Session.getActiveUser().getEmail() || 'system';
    saveSystemSetting('xray_template_config', JSON.stringify(config), user);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ── DRIVE HELPERS ───────────────────────────────────────────
function getDriveFolder_(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

// Removes trailing empty paragraphs and manual page breaks from the body.
function cleanupTrailingBodyWhitespace_(body) {
  if (!body) return;
  try {
    // Strip page breaks anywhere in the body
    var i = 0;
    while (i < body.getNumChildren()) {
      var child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        var para = child.asParagraph();
        for (var j = para.getNumChildren() - 1; j >= 0; j--) {
          var el = para.getChild(j);
          if (el.getType() === DocumentApp.ElementType.PAGE_BREAK) {
            try { el.removeFromParent(); } catch (e) { }
          }
        }
      }
      i++;
    }
    // Drop trailing empty paragraphs
    while (body.getNumChildren() > 1) {
      var lastIdx = body.getNumChildren() - 1;
      var last = body.getChild(lastIdx);
      if (last.getType() !== DocumentApp.ElementType.PARAGRAPH) break;
      var lpara = last.asParagraph();
      if (lpara.getText().trim() !== '') break;
      // Bail if the paragraph holds anything non-text (image, inline drawing…)
      var isPlainEmpty = true;
      for (var k = 0; k < lpara.getNumChildren(); k++) {
        var t = lpara.getChild(k).getType();
        if (t !== DocumentApp.ElementType.TEXT) { isPlainEmpty = false; break; }
      }
      if (!isPlainEmpty) break;
      // Predecessor must be a paragraph for removal to be legal
      var prev = body.getChild(lastIdx - 1);
      if (prev.getType() !== DocumentApp.ElementType.PARAGRAPH) break;
      try { body.removeChild(last); } catch (e) { break; }
    }
  } catch (outerE) {
    // Cleanup is best-effort; never let it block result generation
    Logger.log('cleanupTrailingBodyWhitespace_ ERROR: ' + outerE.message);
  }
}

function getXrayTemplateFile_() {
  var folder = getDriveFolder_(XRAY_TEMPLATE_FOLDER);
  var it = folder.getFilesByName(XRAY_TEMPLATE_FILENAME);
  return it.hasNext() ? it.next() : null;
}

// ── TEMPLATE INFO (for Configurations page) ─────────────────
function getXrayTemplateInfo() {
  var file = getXrayTemplateFile_();
  if (!file) {
    return { found: false, folderName: XRAY_TEMPLATE_FOLDER, filename: XRAY_TEMPLATE_FILENAME };
  }
  return {
    found: true,
    folderName: XRAY_TEMPLATE_FOLDER,
    filename: file.getName(),
    fileId: file.getId(),
    updatedAt: file.getLastUpdated().toISOString()
  };
}

// ── IMAGE PLACEHOLDER HELPER ─────────────────────────────────
function replaceWithImage_(body, placeholder, imageUrl) {
  var found = body.findText(placeholder);
  if (!found) return;
  if (!imageUrl) { body.replaceText(placeholder, ''); return; }
  try {
    // Extract Drive file ID from thumbnail URL or direct URL
    var fileId = null;
    var idMatch = imageUrl.match(/[?&]id=([^&]+)/);
    if (idMatch) {
      fileId = idMatch[1];
    } else {
      var pathMatch = imageUrl.match(/\/d\/([^\/]+)/);
      if (pathMatch) fileId = pathMatch[1];
    }
    var blob = fileId ? DriveApp.getFileById(fileId).getBlob()
      : UrlFetchApp.fetch(imageUrl).getBlob();
    var elem = found.getElement();
    var para = elem.getParent();
    var paraIdx = body.getChildIndex(para);
    var img = body.insertImage(paraIdx, blob);
    // Resize to a reasonable signature size
    img.setWidth(150);
    img.setHeight(60);
    // Remove placeholder paragraph (now shifted +1)
    body.removeChild(body.getChild(paraIdx + 1));
  } catch (e) {
    Logger.log('replaceWithImage_ ERROR [' + placeholder + ']: ' + e.message);
    body.replaceText(placeholder, '');
  }
}

// ── COMPACT TABLE STYLING HELPER ─────────────────────────────
function styleCompactTable_(tbl, rowMeta, NUM_COLS) {
  // Calculate optimal column widths based on content
  var totalPts = 451; // A4 width in points with standard margins

  // Analyze content to determine optimal column widths
  var maxTestLength = 0;
  var maxResultLength = 0;
  var maxUnitLength = 0;
  var maxRefLength = 0;

  for (var r = 1; r < tbl.getNumRows(); r++) {
    var meta = rowMeta[r];
    if (meta.type === 'data') {
      var row = tbl.getRow(r);
      if (row.getNumCells() > 0) maxTestLength = Math.max(maxTestLength, row.getCell(0).getText().length);
      if (row.getNumCells() > 1) maxResultLength = Math.max(maxResultLength, row.getCell(1).getText().length);
      if (row.getNumCells() > 2 && NUM_COLS > 2) maxUnitLength = Math.max(maxUnitLength, row.getCell(2).getText().length);
      if (row.getNumCells() > 3 && NUM_COLS > 3) maxRefLength = Math.max(maxRefLength, row.getCell(3).getText().length);
    }
  }

  // Set column widths based on content
  if (NUM_COLS === 2) {
    var testWidth = Math.min(Math.max(180, maxTestLength * 7), 250);
    var resultWidth = totalPts - testWidth;
    tbl.setColumnWidth(0, testWidth);
    tbl.setColumnWidth(1, resultWidth);
  } else if (NUM_COLS === 3) {
    var testWidth = Math.min(Math.max(150, maxTestLength * 6), 220);
    var unitWidth = Math.min(Math.max(70, maxUnitLength * 6), 100);
    var resultWidth = totalPts - testWidth - unitWidth;
    tbl.setColumnWidth(0, testWidth);
    tbl.setColumnWidth(1, resultWidth);
    tbl.setColumnWidth(2, unitWidth);
  } else {
    var testWidth = Math.min(Math.max(120, maxTestLength * 5), 180);
    var unitWidth = Math.min(Math.max(60, maxUnitLength * 5), 80);
    var refWidth = Math.min(Math.max(90, maxRefLength * 4), 140);
    var resultWidth = totalPts - testWidth - unitWidth - refWidth;
    tbl.setColumnWidth(0, testWidth);
    tbl.setColumnWidth(1, resultWidth);
    tbl.setColumnWidth(2, unitWidth);
    tbl.setColumnWidth(3, refWidth);
  }

  // Style each row and cell
  for (var r = 0; r < tbl.getNumRows(); r++) {
    var row = tbl.getRow(r);
    var meta = rowMeta[r];

    // Set minimal row height
    row.setMinimumHeight(1);

    for (var c = 0; c < NUM_COLS; c++) {
      var cell = row.getCell(c);

      // Minimal cell padding
      cell.setPaddingTop(2);
      cell.setPaddingBottom(2);
      cell.setPaddingLeft(3);
      cell.setPaddingRight(3);

      // Cell background and borders
      cell.setBackgroundColor('#FFFFFF');
      cell.setAttributes({
        [DocumentApp.Attribute.BORDER_COLOR]: '#FFFFFF',
        [DocumentApp.Attribute.BORDER_WIDTH]: 0
      });

      // Compact paragraph spacing inside each cell
      for (var pi = 0; pi < cell.getNumChildren(); pi++) {
        var cp = cell.getChild(pi);
        if (cp.getType() === DocumentApp.ElementType.PARAGRAPH) {
          var cpara = cp.asParagraph();
          cpara.setSpacingBefore(0);
          cpara.setSpacingAfter(0);
          cpara.setLineSpacing(1.0);
          cpara.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        }
      }

      // Text styling
      var txt = cell.editAsText();
      txt.setBold(false).setItalic(false).setFontSize(8);

      if (meta.type === 'thead') {
        txt.setBold(true).setFontSize(8);
      } else if (meta.type === 'header') {
        if (c === 0) txt.setBold(true);
        else cell.setText('');
      } else if (meta.type === 'subheader') {
        if (c === 0) txt.setItalic(true);
        else cell.setText('');
      }
    }
  }

  // Set table spacing
  tbl.setAttributes({
    [DocumentApp.Attribute.SPACING_BEFORE]: 0,
    [DocumentApp.Attribute.SPACING_AFTER]: 0
  });
}

// ── GENERATE RESULT DOCUMENT + SAVE TO PATIENT FOLDER ───────
function generateXrayResult(payload) {
  // payload: { patient, procedure, clinicalData, findings, impression,
  //            encodedBy, branchId, orderId, orderItemId, orderNo, reportDate }
  try {
    // Get Google Doc template ID from System_Settings
    var templateDocId = getSettingValue_('xray_template_doc_id', '');
    if (!templateDocId) {
      return { success: false, message: 'No X-Ray template configured. Go to Global Settings → X-Ray Result Template and paste your Google Doc link.' };
    }

    var p = payload.patient || {};
    var tz = Session.getScriptTimeZone();

    var replacements = {
      '{{PATIENT_NAME}}': p.name || '',
      '{{AGE_SEX}}': (p.age || '') + '/' + (p.sex || ''),
      '{{PATIENT_NO}}': p.patient_no || p.patient_id || '',
      '{{BIRTHDATE}}': p.birthdate || '',
      '{{PHYSICIAN}}': p.physician || '',
      '{{DATE}}': payload.reportDate || Utilities.formatDate(new Date(), tz, 'MMMM dd, yyyy'),
      '{{SERVICE_NAME}}': String(payload.procedure || '').toUpperCase(),
      '{{CLINICAL_DATA}}': payload.clinicalData || ' ',
      '{{FINDINGS}}': payload.findings || ' ',
      '{{IMPRESSION}}': payload.impression || ' ',
      '{{RADTECH_NAME}}': p.radtech_name || '',
      '{{RADTECH_CRED}}': p.radtech_cred || '',
      '{{RADIOLOGIST_NAME}}': p.radiologist_name || '',
      '{{RADIOLOGIST_CRED}}': p.radiologist_cred || ''
    };

    var safeProc = String(payload.procedure || 'RESULT').replace(/[^A-Za-z0-9 _-]/g, '').trim();
    var baseName = (payload.orderNo || payload.orderId || 'XRAY') + ' - ' + safeProc;

    // ── Copy the template and replace placeholders via DocumentApp ──
    var templateFile = DriveApp.getFileById(templateDocId);

    // Check if the template is a Word Document (.docx) instead of a Google Doc
    if (templateFile.getMimeType() !== MimeType.GOOGLE_DOCS) {
      return { success: false, message: 'The template is a Word Document (.docx). It must be a Google Doc! Please open it, click "File > Save as Google Docs", and use the link from the newly converted Google Doc instead.' };
    }

    // Create the copy in the root of the user's Drive temporarily
    var docCopy = templateFile.makeCopy(baseName + '_edit');
    var docDoc = DocumentApp.openById(docCopy.getId());
    var body = docDoc.getBody();

    // Replace in body
    Object.keys(replacements).forEach(function (ph) {
      body.replaceText(ph, replacements[ph] || '');
    });
    // Replace in footer section (if user placed placeholders there)
    var xrayFooter = docDoc.getFooter();
    if (xrayFooter) {
      Object.keys(replacements).forEach(function (ph) {
        xrayFooter.replaceText(ph, replacements[ph] || '');
      });
    }

    // Replace signature image placeholders — check both body and footer
    var xSigC1 = (xrayFooter && xrayFooter.findText('{{RADTECH_SIGNATURE}}')) ? xrayFooter : body;
    replaceWithImage_(xSigC1, '{{RADTECH_SIGNATURE}}', p.radtech_signature_url || '');
    var xSigC2 = (xrayFooter && xrayFooter.findText('{{RADIOLOGIST_SIGNATURE}}')) ? xrayFooter : body;
    replaceWithImage_(xSigC2, '{{RADIOLOGIST_SIGNATURE}}', p.radiologist_signature_url || '');

    cleanupTrailingBodyWhitespace_(body);

    docDoc.saveAndClose();

    // ── Determine target folder (patient sub-folder inside branch root) ──
    var targetFolder = null;
    try {
      var drvCfg = getDriveFolderConfig(payload.branchId);
      var rootId = drvCfg && drvCfg.root_folder_id ? drvCfg.root_folder_id.trim() : '';
      if (rootId) {
        var rootFolder = DriveApp.getFolderById(rootId);
        // Correctly format string to find folder
        var folderName = '';
        if (p.name) folderName = p.name.trim() + ' - ' + (p.patient_id || '').trim();
        else folderName = 'Patient - ' + (p.patient_id || '');

        // Search for existing patient folder, or create it
        var fq = rootFolder.getFoldersByName(folderName);
        if (fq.hasNext()) {
          targetFolder = fq.next();
        } else {
          targetFolder = rootFolder.createFolder(folderName);
        }
      }
    } catch (dfe) {
      Logger.log("generateXrayResult: Target folder finding failed: " + dfe.message);
    }

    if (!targetFolder) {
      targetFolder = getDriveFolder_(XRAY_RESULTS_FOLDER);
    }

    // ── Export PDF → save to patient folder, trash the temp doc ──
    Utilities.sleep(1500); // Give Google Docs a tiny moment to finish syncing the saveAndClose() before PDF export

    var pdfBlob = docCopy.getAs(MimeType.PDF);
    pdfBlob.setName(baseName + '.pdf');
    var pdfFile = targetFolder.createFile(pdfBlob);

    // Trash the temporary Google Doc copy
    docCopy.setTrashed(true);

    return { success: true, pdfId: pdfFile.getId(), pdfUrl: pdfFile.getUrl(), filename: baseName };

  } catch (e) {
    Logger.log('generateXrayResult ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── SAVE X-RAY ENCODING: generate PDF + record in RESULT_ITEMS + mark encoded ──
function saveXrayResultAndPdf(branchId, orderId, orderItemId, servId, servName,
  clinicalData, findings, impression, encodedBy, orderNo) {
  try {
    if (!branchId || !orderId || !orderItemId)
      return { success: false, message: 'Missing required parameters.' };

    // 1. Get patient info from the order
    var ss = getOrderSS_(branchId);
    var ordSh = ss.getSheetByName('LAB_ORDER');
    var itemSh = ss.getSheetByName('LAB_ORDER_ITEM');
    var patient = {};
    if (ordSh && ordSh.getLastRow() >= 2) {
      var oCols = Math.max(ordSh.getLastColumn(), 20);
      var ordRow = ordSh.getRange(2, 1, ordSh.getLastRow() - 1, oCols).getValues()
        .find(function (r) { return String(r[0]).trim() === orderId; });
      if (ordRow) {
        patient.name = String(ordRow[11] || ordRow[3] || '').trim();
        patient.patient_id = String(ordRow[3] || '').trim();
        patient.physician = String(ordRow[12] || '').trim();
        // parse age/sex from patient_name if needed; pull from Patients sheet
        try {
          var masterSS = getSS_();
          var patSh = masterSS.getSheetByName('Patients_' + branchId) ||
            masterSS.getSheetByName('Patients');
          if (!patSh) {
            var bss = getOrderSS_(branchId);
            patSh = bss.getSheetByName('Patients');
          }
          if (patSh && patSh.getLastRow() >= 2) {
            var pRow = patSh.getRange(2, 1, patSh.getLastRow() - 1, 8).getValues()
              .find(function (r) { return String(r[0]).trim() === patient.patient_id; });
            if (pRow) {
              patient.name = String(pRow[1]).trim() + ', ' + String(pRow[2]).trim();
              patient.sex = String(pRow[4] || '').trim();
              var dob = pRow[5] ? new Date(pRow[5]) : null;
              patient.birthdate = dob ? Utilities.formatDate(dob, Session.getScriptTimeZone(), 'MMMM dd, yyyy') : '';
              patient.age = dob ? Math.floor((new Date() - dob) / (365.25 * 24 * 3600 * 1000)) + '' : '';
            }
          }
        } catch (pe) { /* non-fatal */ }
      }
    }

    // 2. Get tech's own name/credentials/signature from their account record
    try {
      var techInfo = getTechInfo(encodedBy);
      if (techInfo && techInfo.success) {
        patient.radtech_name = techInfo.name || '';
        patient.radtech_cred = techInfo.credentials || '';
        patient.radtech_signature_url = techInfo.signature_url || '';
      }
    } catch (se) { /* non-fatal */ }

    // 3. Get radiologist (branch-level) from Branch Settings
    try {
      var xraySig = getXraySignatures(branchId);
      if (xraySig && xraySig.success) {
        patient.radiologist_name = xraySig.radiologist ? xraySig.radiologist.name : '';
        patient.radiologist_cred = xraySig.radiologist ? xraySig.radiologist.credentials : '';
        patient.radiologist_signature_url = xraySig.radiologist ? (xraySig.radiologist.signature_url || '') : '';
      }
    } catch (se) { /* non-fatal */ }

    // 3. Generate PDF
    var tz = Session.getScriptTimeZone();
    var gen = generateXrayResult({
      patient: patient,
      procedure: servName || servId,
      clinicalData: clinicalData || '',
      findings: findings || '',
      impression: impression || '',
      encodedBy: encodedBy,
      branchId: branchId,
      orderId: orderId,
      orderItemId: orderItemId,
      orderNo: orderNo || orderId,
      reportDate: Utilities.formatDate(new Date(), tz, 'MMMM dd, yyyy')
    });
    if (!gen.success) return { success: false, message: gen.message || 'PDF generation failed.' };

    // 4. Save to RESULT_ITEMS (unit=XRAY_PDF so UI can distinguish)
    var itemsSh = _getResultItemsSheet_(ss);
    var lr = itemsSh.getLastRow();
    var existIdx = -1;
    if (lr >= 2) {
      existIdx = itemsSh.getRange(2, 1, lr - 1, 3).getValues()
        .findIndex(function (r) { return String(r[1]).trim() === orderId && String(r[2]).trim() === orderItemId; });
    }
    if (existIdx !== -1) {
      // cols 5-14: pdfUrl, type, '', '', encodedBy, timestamp, result_type, clinicalData, findings, impression
      itemsSh.getRange(existIdx + 2, 5, 1, 10).setValues([[gen.pdfUrl, 'XRAY_PDF', '', '', encodedBy, new Date(), 'xray', clinicalData || '', findings || '', impression || '']]);
    } else {
      var rid = 'RI-' + Math.random().toString(16).substr(2, 8).toUpperCase();
      itemsSh.appendRow([rid, orderId, orderItemId, servName || 'X-Ray Result', gen.pdfUrl, 'XRAY_PDF', '', '', encodedBy, new Date(), 'xray', clinicalData || '', findings || '', impression || '']);
    }

    // 5. Mark item checkpoint as encoded
    if (itemSh && itemSh.getLastRow() >= 2) {
      ensureItemCols_(itemSh);
      var iRows = itemSh.getRange(2, 1, itemSh.getLastRow() - 1, 2).getValues();
      var iIdx = iRows.findIndex(function (r) { return String(r[0]).trim() === orderItemId && String(r[1]).trim() === orderId; });
      if (iIdx !== -1) itemSh.getRange(iIdx + 2, 16).setValue(new Date()); // encoded_at col 16
    }

    // 6. Auto-advance order status
    var progress = _checkOrderProgress_(ss, itemSh, orderId, branchId);

    return { success: true, pdfUrl: gen.pdfUrl, order_status: progress.newStatus };
  } catch (e) {
    Logger.log('saveXrayResultAndPdf ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
//  LAB RESULT PDF GENERATION
//  Uses a single Google Doc template (lab_template_doc_id setting).
//  Replaces patient placeholders and inserts a results table
//  where {{RESULTS_TABLE}} appears in the template.
//  Each service call produces one separate PDF.
// ════════════════════════════════════════════════════════════════

function generateLabResult(payload) {
  // payload: { patient, serviceName, params:[{param_name,result_value,unit,reference_range,remarks,field_type}],
  //            encodedBy, branchId, orderId, orderItemId, orderNo, reportDate }
  try {
    var templateDocId = getSettingValue_('lab_template_doc_id', '');
    if (!templateDocId) {
      return { success: false, message: 'No Lab Result template configured. Go to Global Settings → Lab Result Template and paste your Google Doc link.' };
    }

    var templateFile = DriveApp.getFileById(templateDocId);
    if (templateFile.getMimeType() !== MimeType.GOOGLE_DOCS) {
      return { success: false, message: 'The lab template must be a Google Doc (not .docx). Open it, click File > Save as Google Docs, then use that link.' };
    }

    var p = payload.patient || {};
    var tz = Session.getScriptTimeZone();
    var now = payload.reportDate || Utilities.formatDate(new Date(), tz, 'MMMM dd, yyyy');

    var replacements = {
      '{{PATIENT_NAME}}': p.name || '',
      '{{AGE_SEX}}': (p.age || '') + (p.sex ? '/' + p.sex : ''),
      '{{PATIENT_NO}}': p.patient_no || p.patient_id || '',
      '{{BIRTHDATE}}': p.birthdate || '',
      '{{PHYSICIAN}}': p.physician || '',
      '{{DATE}}': now,
      '{{SERVICE_NAME}}': String(payload.serviceName || '').toUpperCase(),
      '{{MEDTECH_NAME}}': p.medtech_name || '',
      '{{MEDTECH_CRED}}': p.medtech_cred || '',
      '{{MEDTECH_LICENSE}}': p.medtech_license_no || '',
      '{{PATHOLOGIST_NAME}}': p.pathologist_name || '',
      '{{PATHOLOGIST_CRED}}': p.pathologist_cred || '',
      '{{PATHOLOGIST_LICENSE}}': p.pathologist_license_no || ''
    };


    var safeService = String(payload.serviceName || 'LAB').replace(/[^A-Za-z0-9 _-]/g, '').trim();
    var baseName = (payload.orderNo || payload.orderId || 'LAB') + ' - ' + safeService;

    // Clone template
    var docCopy = templateFile.makeCopy(baseName + '_edit');
    var docDoc = DocumentApp.openById(docCopy.getId());
    var body = docDoc.getBody();

    // Replace scalar placeholders
    Object.keys(replacements).forEach(function (ph) {
      body.replaceText(ph, replacements[ph] || '');
    });
    // Replace in footer section (if user placed placeholders there)
    var footer = docDoc.getFooter();
    if (footer) {
      Object.keys(replacements).forEach(function (ph) {
        footer.replaceText(ph, replacements[ph] || '');
      });
    }

    // Find and replace {{RESULTS_TABLE}} with an actual table
    var params = payload.params || [];
    var numChildren = body.getNumChildren();
    var tableIdx = -1;
    for (var i = 0; i < numChildren; i++) {
      var child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH &&
        child.asText().getText().indexOf('{{RESULTS_TABLE}}') !== -1) {
        tableIdx = i;
        break;
      }
    }

    if (tableIdx !== -1) {
      // ── Dynamic column detection ──────────────────────────────
      var dataParams = params.filter(function (p) {
        var ft = p.field_type || 'numeric';
        return ft !== 'header' && ft !== 'subheader';
      });
      var hasUnit = dataParams.some(function (p) { return (p.unit || '').trim() !== ''; });
      var hasRefRange = dataParams.some(function (p) { return (p.reference_range || '').trim() !== ''; });

      var colHeaders = ['TEST', 'RESULT'];
      if (hasUnit) colHeaders.push('UNIT');
      if (hasRefRange) colHeaders.push('REFERENCE RANGE');
      var NUM_COLS = colHeaders.length;

      // ── Build row data ────────────────────────────────────────
      var allRows = [colHeaders];
      var rowMeta = [{ type: 'thead' }];

      params.forEach(function (p) {
        var ft = p.field_type || 'numeric';
        if (ft === 'header' || ft === 'subheader') {
          var label = (p.param_name || '').toUpperCase();
          var sRow = [label];
          for (var i = 1; i < NUM_COLS; i++) sRow.push('');
          allRows.push(sRow);
          rowMeta.push({ type: ft });
        } else {
          var dRow = [p.param_name || '', p.result_value || ''];
          if (hasUnit) dRow.push(p.unit || '');
          if (hasRefRange) dRow.push(p.reference_range || '');
          allRows.push(dRow);
          rowMeta.push({ type: 'data' });
        }
      });

      body.removeChild(body.getChild(tableIdx));
      var tbl = body.insertTable(tableIdx, allRows);

      // ── Apply compact table styling ─────────────────────────
      styleCompactTable_(tbl, rowMeta, NUM_COLS);

      // ── Remove any "Test" text that might remain ─────────────
      var tablePos = body.getChildIndex(tbl);
      if (tablePos + 1 < body.getNumChildren()) {
        var nextElement = body.getChild(tablePos + 1);
        if (nextElement && nextElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
          var text = nextElement.asParagraph().getText().trim();
          if (text === 'Test' || text === '') {
            body.removeChild(nextElement);
          }
        }
      }
    }

    // Replace signature image placeholders — check both body and footer
    var sigContainer = (footer && footer.findText('{{MEDTECH_SIGNATURE}}')) ? footer : body;
    replaceWithImage_(sigContainer, '{{MEDTECH_SIGNATURE}}', p.medtech_signature_url || '');
    var sigContainer2 = (footer && footer.findText('{{PATHOLOGIST_SIGNATURE}}')) ? footer : body;
    replaceWithImage_(sigContainer2, '{{PATHOLOGIST_SIGNATURE}}', p.pathologist_signature_url || '');

    // Trim trailing blank paragraphs / manual page breaks in the body.
    cleanupTrailingBodyWhitespace_(body);

    docDoc.saveAndClose();

    // Determine patient folder in Drive
    var targetFolder = null;
    try {
      var drvCfg = getDriveFolderConfig(payload.branchId);
      var rootId = drvCfg && drvCfg.root_folder_id ? drvCfg.root_folder_id.trim() : '';
      if (rootId) {
        var rootFolder = DriveApp.getFolderById(rootId);
        var folderName = p.name ? p.name.trim() + ' - ' + (p.patient_id || '').trim()
          : 'Patient - ' + (p.patient_id || '');
        var fq = rootFolder.getFoldersByName(folderName);
        targetFolder = fq.hasNext() ? fq.next() : rootFolder.createFolder(folderName);
      }
    } catch (dfe) { Logger.log('generateLabResult: folder error: ' + dfe.message); }

    if (!targetFolder) targetFolder = getDriveFolder_('A-Lab Results');

    Utilities.sleep(1500);
    var pdfBlob = docCopy.getAs(MimeType.PDF);
    pdfBlob.setName(baseName + '.pdf');
    var pdfFile = targetFolder.createFile(pdfBlob);
    docCopy.setTrashed(true);

    return { success: true, pdfId: pdfFile.getId(), pdfUrl: pdfFile.getUrl(), filename: baseName };
  } catch (e) {
    Logger.log('generateLabResult ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── SAVE LAB ENCODING: fill form → generate PDF → record in RESULT_ITEMS ──
function saveLabResultAndPdf(branchId, orderId, orderItemId, servId, servName, params, encodedBy, orderNo) {
  try {
    if (!branchId || !orderId || !orderItemId)
      return { success: false, message: 'Missing required parameters.' };

    var ss = getOrderSS_(branchId);
    var ordSh = ss.getSheetByName('LAB_ORDER');
    var itemSh = ss.getSheetByName('LAB_ORDER_ITEM');
    var patient = {};

    // Pull patient info from the order
    if (ordSh && ordSh.getLastRow() >= 2) {
      var oCols = Math.max(ordSh.getLastColumn(), 20);
      var ordRow = ordSh.getRange(2, 1, ordSh.getLastRow() - 1, oCols).getValues()
        .find(function (r) { return String(r[0]).trim() === orderId; });
      if (ordRow) {
        patient.name = String(ordRow[11] || ordRow[3] || '').trim();
        patient.patient_id = String(ordRow[3] || '').trim();
        patient.physician = String(ordRow[12] || '').trim();
        try {
          var masterSS = getSS_();
          var patSh = masterSS.getSheetByName('Patients_' + branchId) ||
            masterSS.getSheetByName('Patients');
          if (!patSh) patSh = getOrderSS_(branchId).getSheetByName('Patients');
          if (patSh && patSh.getLastRow() >= 2) {
            var pRow = patSh.getRange(2, 1, patSh.getLastRow() - 1, 8).getValues()
              .find(function (r) { return String(r[0]).trim() === patient.patient_id; });
            if (pRow) {
              patient.name = String(pRow[1]).trim() + ', ' + String(pRow[2]).trim();
              patient.sex = String(pRow[4] || '').trim();
              var dob = pRow[5] ? new Date(pRow[5]) : null;
              patient.birthdate = dob ? Utilities.formatDate(dob, Session.getScriptTimeZone(), 'MMMM dd, yyyy') : '';
              patient.age = dob ? Math.floor((new Date() - dob) / (365.25 * 24 * 3600 * 1000)) + '' : '';
            }
          }
        } catch (pe) { /* non-fatal */ }
      }
    }

    // Fetch tech's own name/credentials/signature from their account record
    try {
      var techInfo = getTechInfo(encodedBy);
      if (techInfo && techInfo.success) {
        patient.medtech_name = techInfo.name || '';
        patient.medtech_cred = techInfo.credentials || '';
        patient.medtech_license_no = techInfo.license_no || '';
        patient.medtech_signature_url = techInfo.signature_url || '';
      }
    } catch (se) { /* non-fatal */ }

    // Fetch pathologist (branch-level) from branch settings
    try {
      var labSig = getLabSignatures(branchId);
      if (labSig && labSig.success) {
        patient.pathologist_name = labSig.pathologist ? labSig.pathologist.name : '';
        patient.pathologist_cred = labSig.pathologist ? labSig.pathologist.credentials : '';
        patient.pathologist_license_no = labSig.pathologist ? (labSig.pathologist.license_no || '') : '';
        patient.pathologist_signature_url = labSig.pathologist ? (labSig.pathologist.signature_url || '') : '';
      }
    } catch (se) { /* non-fatal */ }

    // Generate PDF
    var tz = Session.getScriptTimeZone();
    var gen = generateLabResult({
      patient: patient,
      serviceName: servName || servId,
      params: params || [],
      encodedBy: encodedBy,
      branchId: branchId,
      orderId: orderId,
      orderItemId: orderItemId,
      orderNo: orderNo || orderId,
      reportDate: Utilities.formatDate(new Date(), tz, 'MMMM dd, yyyy')
    });
    if (!gen.success) return { success: false, message: gen.message || 'PDF generation failed.' };

    // Save PDF URL to RESULT_ITEMS (unit=LAB_PDF)
    var itemsSh = _getResultItemsSheet_(ss);
    var lr = itemsSh.getLastRow();
    var existIdx = -1;
    if (lr >= 2) {
      existIdx = itemsSh.getRange(2, 1, lr - 1, 3).getValues()
        .findIndex(function (r) { return String(r[1]).trim() === orderId && String(r[2]).trim() === orderItemId; });
    }
    if (existIdx !== -1) {
      itemsSh.getRange(existIdx + 2, 5, 1, 6).setValues([[gen.pdfUrl, 'LAB_PDF', '', '', encodedBy, new Date()]]);
    } else {
      var rid = 'RI-' + Math.random().toString(16).substr(2, 8).toUpperCase();
      itemsSh.appendRow([rid, orderId, orderItemId, servName || 'Lab Result', gen.pdfUrl, 'LAB_PDF', '', '', encodedBy, new Date()]);
    }

    // Mark item as encoded
    if (itemSh && itemSh.getLastRow() >= 2) {
      ensureItemCols_(itemSh);
      var iRows = itemSh.getRange(2, 1, itemSh.getLastRow() - 1, 2).getValues();
      var iIdx = iRows.findIndex(function (r) { return String(r[0]).trim() === orderItemId && String(r[1]).trim() === orderId; });
      if (iIdx !== -1) itemSh.getRange(iIdx + 2, 16).setValue(new Date());
    }

    var progress = _checkOrderProgress_(ss, itemSh, orderId, branchId);
    return { success: true, pdfUrl: gen.pdfUrl, order_status: progress.newStatus };
  } catch (e) {
    Logger.log('saveLabResultAndPdf ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── DELETE ITEM RESULT ──────────────────────────────────────────
// Trashes the existing PDF from Drive, clears the RESULT_ITEMS row,
// and resets encoded_at on LAB_ORDER_ITEM so the tech can re-encode.
// Returns previous xray fields for pre-fill on re-edit.
function deleteItemResult(branchId, orderId, orderItemId) {
  try {
    if (!branchId || !orderId || !orderItemId)
      return { success: false, message: 'Missing parameters.' };

    var ss = getOrderSS_(branchId);
    var rSh = _getResultItemsSheet_(ss);

    if (rSh.getLastRow() < 2) return { success: false, message: 'No result found.' };

    var numCols = Math.max(rSh.getLastColumn(), 14);
    var rows = rSh.getRange(2, 1, rSh.getLastRow() - 1, numCols).getValues();
    var idx = rows.findIndex(function (r) {
      return String(r[1]).trim() === orderId && String(r[2]).trim() === orderItemId;
    });
    if (idx === -1) return { success: false, message: 'Result not found.' };

    var row = rows[idx];
    var pdfUrl = String(row[4] || '').trim();
    var prevCD = String(row[11] || '').trim();
    var prevFN = String(row[12] || '').trim();
    var prevIM = String(row[13] || '').trim();

    // Trash the PDF file from Drive
    if (pdfUrl) {
      try {
        var idMatch = pdfUrl.match(/(?:\/d\/|[?&]id=)([-\w]{10,})/);
        if (idMatch) DriveApp.getFileById(idMatch[1]).setTrashed(true);
      } catch (fe) { Logger.log('deleteItemResult: could not trash PDF: ' + fe.message); }
    }

    // Clear PDF URL and unit from RESULT_ITEMS (keep row identity)
    var clearCols = [];
    for (var i = 0; i < 10; i++) clearCols.push('');
    rSh.getRange(idx + 2, 5, 1, 10).setValues([clearCols]);

    // Reset encoded_at (col 16) on LAB_ORDER_ITEM
    var itemSh = ss.getSheetByName('LAB_ORDER_ITEM');
    if (itemSh && itemSh.getLastRow() >= 2) {
      var iRows = itemSh.getRange(2, 1, itemSh.getLastRow() - 1, 2).getValues();
      var iIdx = iRows.findIndex(function (r) {
        return String(r[0]).trim() === orderItemId && String(r[1]).trim() === orderId;
      });
      if (iIdx !== -1) itemSh.getRange(iIdx + 2, 16).setValue('');
    }

    writeAuditLog_('RESULT_DELETED', { branch_id: branchId, order_id: orderId, item_id: orderItemId });
    return { success: true, clinical_data: prevCD, findings: prevFN, impression: prevIM };
  } catch (e) {
    Logger.log('deleteItemResult ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}
