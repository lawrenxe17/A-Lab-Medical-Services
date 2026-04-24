function processBulkImportPackages(payload) {
  try {
    if (!payload || !payload.length) return { success: false, message: 'No data provided.' };

    const pkgSh   = getPkgSheet_();
    const itemsSh = getPkgItemsSheet_();
    const servSh  = getLabServSheet_();
    const now     = new Date();
    const stats   = { created: 0, skipped: 0, servs: 0 };

    // Build existing package name map for duplicate check
    const lr = pkgSh.getLastRow();
    const existingNames = lr >= 2
      ? pkgSh.getRange(2, 2, lr - 1, 1).getValues().flat().map(v => String(v).trim().toLowerCase())
      : [];

    payload.forEach(pkg => {
      const name    = (pkg.package_name || '').trim();
      const fee     = pkg.default_fee || 0;
      const servIds = (pkg.serv_ids || []).filter(id => id && String(id).trim());
      const unmatched = pkg.unmatched || [];

      if (!name) return;

      // Skip duplicates
      if (existingNames.includes(name.toLowerCase())) {
        stats.skipped++;
        return;
      }

      // If no matched AND no unmatched, skip
      if (!servIds.length && !unmatched.length) {
        stats.skipped++;
        return;
      }

      const pkgId = 'PKG-' + Math.random().toString(16).substr(2, 8).toUpperCase();
      pkgSh.appendRow([pkgId, name, (pkg.description || '').trim(), fee, 1, now, now]);

      // Add matched services to package
      servIds.forEach(sid => {
        const itemId = 'PKGI-' + Math.random().toString(16).substr(2, 8).toUpperCase();
        itemsSh.appendRow([itemId, pkgId, sid.trim(), now]);
        stats.servs++;
      });

      // Handle unmatched services: create them in Lab_Services and add them to the package
      unmatched.forEach(rawName => {
        if (!rawName.trim()) return;
        
        const newServId = 'SRV-' + Math.random().toString(16).substr(2, 8).toUpperCase();
        
        // Append to Lab_Services
        // A=serv_id, B=cat_id, C=serv_name, D=default_fee, E=specimen_type, F=is_ph_covered,
        // G=is_active, H=created_at, I=updated_at, J=service_type, K=is_consultation, L=template_url
        servSh.appendRow([
          newServId, '', rawName.trim(), 0, '', 0, 1, now, now, '', 0, ''
        ]);
        
        // Append to Package items
        const newItemId = 'PKGI-' + Math.random().toString(16).substr(2, 8).toUpperCase();
        itemsSh.appendRow([newItemId, pkgId, newServId, now]);
        stats.servs++;
      });

      propagatePkgToAllBranches_(pkgId);
      existingNames.push(name.toLowerCase());
      stats.created++;
    });

    Logger.log('processBulkImportPackages: ' + stats.created + ' created, ' + stats.skipped + ' skipped.');
    return { success: true, stats: stats };
  } catch(err) {
    Logger.log('processBulkImportPackages ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}

function getBulkImportExistingData() {
  try {
    const deptSheet = getDeptSheet_();
    const catSheet  = getCatSheet_();
    const servSheet = getLabServSheet_();

    const deptRows = deptSheet.getLastRow() >= 2
      ? deptSheet.getRange(2, 1, deptSheet.getLastRow() - 1, 6).getValues() : [];
    const catRows  = catSheet.getLastRow() >= 2
      ? catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 6).getValues() : [];
    const servRows = servSheet.getLastRow() >= 2
      ? servSheet.getRange(2, 1, servSheet.getLastRow() - 1, 11).getValues() : [];

    return {
      depts: deptRows.map(r => ({ id: String(r[0]), name: String(r[1]) })),
      cats:  catRows.map(r  => ({ id: String(r[0]), dept_id: String(r[1]), name: String(r[2]) })),
      servs: servRows.map(r => ({ id: String(r[0]), cat_id:  String(r[1]), name: String(r[2]) }))
    };
  } catch(err) {
    Logger.log('getBulkImportExistingData ERROR: ' + err.message);
    return { depts: [], cats: [], servs: [] };
  }
}

function processBulkImport(payload) {
  try {
    if (!payload || !payload.length) return { success: false, message: 'No data provided.' };

    const ssMap = {
      dept: getDeptSheet_(),
      cat:  getCatSheet_(),
      serv: getLabServSheet_()
    };

    const deptRows = ssMap.dept.getLastRow() >= 2
      ? ssMap.dept.getRange(2, 1, ssMap.dept.getLastRow() - 1, 9).getValues() : [];
    const catRows  = ssMap.cat.getLastRow() >= 2
      ? ssMap.cat.getRange(2, 1, ssMap.cat.getLastRow() - 1, 9).getValues() : [];
    const servRows = ssMap.serv.getLastRow() >= 2
      ? ssMap.serv.getRange(2, 1, ssMap.serv.getLastRow() - 1, 11).getValues() : [];

    const now = new Date();
    const stats = { depts: 0, cats: 0, servs: 0 };

    const existing = {
      depts: Object.fromEntries(deptRows.map(r => [String(r[1]).trim().toLowerCase(), String(r[0]).trim()])),
      cats:  Object.fromEntries(catRows.map(r  => [String(r[1]).trim() + '_' + String(r[2]).trim().toLowerCase(), String(r[0]).trim()])),
      servs: Object.fromEntries(servRows.map(r => [String(r[1]).trim() + '_' + String(r[2]).trim().toLowerCase(), String(r[0]).trim()]))
    };

    const newDepts = [];
    const newCats  = [];
    const newServs = [];

    payload.forEach(row => {
      const deptName  = (row.dept_name  || '').trim();
      const deptClass = (row.dept_class || 'lab').trim();
      const catName   = (row.cat_name   || '').trim();
      const servName  = (row.serv_name  || '').trim();
      const fee       = row.fee || 0;
      const isPhil    = row.is_philhealth_covered ? 1 : 0;

      if (!deptName || !catName || !servName) return;

      // Ensure Department
      let deptId = existing.depts[deptName.toLowerCase()];
      if (!deptId) {
        deptId = 'DEPT-' + Math.random().toString(16).substr(2, 8).toUpperCase();
        existing.depts[deptName.toLowerCase()] = deptId;
        newDepts.push([deptId, deptName, 1, now, now, deptClass]);
        stats.depts++;
      }

      // Ensure Category
      const catKey = deptId + '_' + catName.toLowerCase();
      let catId = existing.cats[catKey];
      if (!catId) {
        catId = 'CAT-' + Math.random().toString(16).substr(2, 8).toUpperCase();
        existing.cats[catKey] = catId;
        newCats.push([catId, deptId, catName, 1, now, now]);
        stats.cats++;
      }

      // Ensure Service (skip if exists)
      const servKey = catId + '_' + servName.toLowerCase();
      if (!existing.servs[servKey]) {
        const servId = 'SERV-' + Math.random().toString(16).substr(2, 8).toUpperCase();
        existing.servs[servKey] = servId;
        const isCon = deptClass === 'consultation' ? 1 : 0;
        // Cols: serv_id, cat_id, serv_name, default_fee, specimen_type, is_philhealth_covered, is_active, created_at, updated_at, service_type, is_consultation
        newServs.push([servId, catId, servName, fee, '', isPhil, 1, now, now, deptClass, isCon]);
        stats.servs++;
      }
    });

    // Commit
    if (newDepts.length) ssMap.dept.getRange(ssMap.dept.getLastRow() + 1, 1, newDepts.length, newDepts[0].length).setValues(newDepts);
    if (newCats.length)  ssMap.cat.getRange(ssMap.cat.getLastRow()   + 1, 1, newCats.length,  newCats[0].length).setValues(newCats);
    if (newServs.length) ssMap.serv.getRange(ssMap.serv.getLastRow() + 1, 1, newServs.length, newServs[0].length).setValues(newServs);

    // Propagate to all branches
    newDepts.forEach(d => propagateDeptToAllBranches_(d[0]));
    newCats.forEach(c  => propagateCatToAllBranches_(c[0]));
    newServs.forEach(s => propagateServToAllBranches_(s[0]));

    Logger.log('Bulk Import: ' + stats.servs + ' services, ' + stats.cats + ' categories, ' + stats.depts + ' departments.');
    return { success: true, count: payload.length, stats: stats };
  } catch(err) {
    Logger.log('processBulkImport ERROR: ' + err.message);
    return { success: false, message: err.message };
  }
}
