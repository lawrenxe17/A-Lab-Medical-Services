function upgradeDatabaseSchema() {
  try {
    const ss = getSS_();
    let logs = [];

    // 1. Upgrade Departments Sheet
    const deptSh = ss.getSheetByName('Departments');
    if (deptSh) {
      let lr = deptSh.getLastRow();
      let lc = Math.max(deptSh.getLastColumn(), 5);
      let headers = deptSh.getRange(1, 1, 1, lc).getValues()[0];
      
      if (headers.indexOf('department_type') === -1) {
        deptSh.getRange(1, 6).setValue('department_type').setFontWeight('bold').setBackground('#0060b0').setFontColor('#fff');
        if (lr >= 2) {
          const data = deptSh.getRange(2, 1, lr - 1, 6).getValues();
          data.forEach(r => {
            if (!r[5]) {
              const name = String(r[1]).toLowerCase();
              r[5] = (name.includes('x-ray') || name.includes('xray') || name.includes('radiolog')) ? 'xray' : 'lab';
            }
          });
          deptSh.getRange(2, 1, lr - 1, 6).setValues(data);
        }
        logs.push('Added "department_type" to Departments sheet.');
      } else {
        logs.push('Departments sheet is already up to date.');
      }
    }

    // 2. Upgrade Technologists Sheet
    const techSh = ss.getSheetByName('Technologists');
    if (techSh) {
      let lr = techSh.getLastRow();
      let lc = Math.max(techSh.getLastColumn(), 13);
      let headers = techSh.getRange(1, 1, 1, lc).getValues()[0];

      if (headers.indexOf('assigned_deps') === -1 && lc < 14) {
        techSh.getRange(1, 14).setValue('assigned_deps').setFontWeight('bold').setBackground('#0060b0').setFontColor('#fff');
        if (lr >= 2) {
          const data = techSh.getRange(2, 10, lr - 1, 5).getValues(); // J(10) to N(14)
          data.forEach(r => {
            const role = String(r[0]).toLowerCase(); // Role is col J (idx 0 relative)
            if (!r[4]) {
              r[4] = role.includes('radiolog') ? 'xray' : 'lab';
            }
          });
          techSh.getRange(2, 10, lr - 1, 5).setValues(data);
        }
        logs.push('Added "assigned_deps" to Technologists sheet.');
      } else {
         logs.push('Technologists sheet is already up to date.');
      }
    }

    return { success: true, message: 'Database schema successfully upgraded!\n\n' + logs.join('\n') };
  } catch(e) {
    Logger.log('upgradeDatabaseSchema ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}