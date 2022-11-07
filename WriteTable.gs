/**
 * Copy and paste BRG 8 DENGGI.
 * @param {Object} projectVar instance of getProjectVariable()
 * @param {Object} daerahList pass daerah object in a list
 * @param {Object} targetTable instance of Table (Google Docs)
 */
function fetchDengueData(projectVar, daerahList, targetTable) {
  let sourceRange = projectVar.sheetDenggi.getRange('A14:Q25');
  let sourceValues = sourceRange.getValues();
  daerahList.forEach(daerah => {
    let values = sourceValues[daerah.position];
    values.splice(2,1); // Clean data
    // Think of the `appendTableRow()` as <tr> in HTML where you have to set the cell using <td> which is `appendTableCell()`
    let newTableRow = targetTable.appendTableRow();
    values.forEach(item => {
      if (typeof item == 'number') {
        newTableRow.appendTableCell(item.toFixed());
      } else {
        newTableRow.appendTableCell(item);
      }
    });
  });
  let sumValues = sourceValues[11];
  sumValues.splice(2,1); // Clean data
  let sumTableRow = targetTable.appendTableRow();
  sumValues.forEach(item => {
    if (typeof item == 'number') {
      sumTableRow.appendTableCell(item.toFixed());
    } else {
      sumTableRow.appendTableCell(item);
    }
  });
  targetTable.setBorderWidth(0.75);
}

/**
 * Copy and paste BRG 9 LALAT.
 * @param {Object} projectVar instance of getProjectVariable()
 * @param {Object} daerahList pass daerah object in a list
 * @param {Object} targetTable instance of Table (Google Docs)
 */
function fetchLalatData(projectVar, daerahList, targetTable) {
  let sourceRange = projectVar.sheetLalat.getRange('A14:O25');
  let sourceValues = sourceRange.getValues();
  daerahList.forEach(daerah => {
    let values = sourceValues[daerah.position];
    // Think of the `appendTableRow()` as <tr> in HTML where you have to set the cell using <td> which is `appendTableCell()`
    let newTableRow = targetTable.appendTableRow();
    values.forEach(item => {
      if (typeof item == 'number') {
        newTableRow.appendTableCell(item.toFixed());
      } else {
        newTableRow.appendTableCell(item);
      }
    });
  });
  let sumValues = sourceValues[11];
  let sumTableRow = targetTable.appendTableRow();
  sumValues.forEach(item => {
    if (typeof item == 'number') {
      sumTableRow.appendTableCell(item.toFixed());
    } else {
      sumTableRow.appendTableCell(item);
    }
  });
  targetTable.setBorderWidth(0.75);
}

/**
 * Copy and paste BRG 10 LITI.
 * @param {Object} projectVar instance of getProjectVariable()
 * @param {Object} daerahList pass daerah object in a list
 * @param {Object} targetTable instance of Table (Google Docs)
 */
function fetchLiTiData(projectVar, daerahList, targetTable) {
  let sourceRange = projectVar.sheetLiTi.getRange('A14:O25');
  let sourceValues = sourceRange.getValues();
  daerahList.forEach(daerah => {
    let values = sourceValues[daerah.position];
    // Think of the `appendTableRow()` as <tr> in HTML where you have to set the cell using <td> which is `appendTableCell()`
    let newTableRow = targetTable.appendTableRow();
    values.forEach(item => {
      if (typeof item == 'number') {
        newTableRow.appendTableCell(item.toFixed());
      } else {
        newTableRow.appendTableCell(item);
      }
    });
  });
  let sumValues = sourceValues[11];
  let sumTableRow = targetTable.appendTableRow();
  sumValues.forEach(item => {
    if (typeof item == 'number') {
      sumTableRow.appendTableCell(item.toFixed());
    } else {
      sumTableRow.appendTableCell(item);
    }
  });
  targetTable.setBorderWidth(0.75);
}

/**
 * Copy and paste BRG 5.1 HECC.
 * @param {Object} projectVar instance of getProjectVariable()
 * @param {Object} daerahList pass daerah object in a list
 * @param {Object} targetTable instance of Table (Google Docs)
 */
function fetchHeccOne(projectVar, daerahList, targetTable) {
  let sourceRange = projectVar.sheetHeccOne.getRange('B14:X25');
  let sourceValues = sourceRange.getValues();
  daerahList.forEach(daerah => {
    let values = sourceValues[daerah.position];
    // Think of the `appendTableRow()` as <tr> in HTML where you have to set the cell using <td> which is `appendTableCell()`
    let newTableRow = targetTable.appendTableRow();
    values.forEach(item => {
      if (typeof item == 'number') {
        newTableRow.appendTableCell(item.toFixed());
      } else {
        newTableRow.appendTableCell(item);
      }
    });
  });
  let sumValues = sourceValues[11];
  let sumTableRow = targetTable.appendTableRow();
  sumValues.forEach(item => {
    if (typeof item == 'number') {
      sumTableRow.appendTableCell(item.toFixed());
    } else {
      sumTableRow.appendTableCell(item);
    }
  });
  targetTable.setBorderWidth(0.75);
}

/**
 * Copy and paste BRG 5.2 HECC.
 * @param {Object} projectVar instance of getProjectVariable()
 * @param {Object} daerahList pass daerah object in a list
 * @param {Object} targetTable instance of Table (Google Docs)
 */
function fetchHeccTwo(projectVar, daerahList, targetTable) {
  let sourceRange = projectVar.sheetHeccTwo.getRange('B14:V25');
  let sourceValues = sourceRange.getValues();
  daerahList.forEach(daerah => {
    let values = sourceValues[daerah.position];
    // Think of the `appendTableRow()` as <tr> in HTML where you have to set the cell using <td> which is `appendTableCell()`
    let newTableRow = targetTable.appendTableRow();
    values.forEach(item => {
      if (typeof item == 'number') {
        newTableRow.appendTableCell(item.toFixed());
      } else {
        newTableRow.appendTableCell(item);
      }
    });
  });
  let sumValues = sourceValues[11];
  let sumTableRow = targetTable.appendTableRow();
  sumValues.forEach(item => {
    if (typeof item == 'number') {
      sumTableRow.appendTableCell(item.toFixed());
    } else {
      sumTableRow.appendTableCell(item);
    }
  });
  targetTable.setBorderWidth(0.75);
}

/**
 * Copy and paste BRG 17 AKTIVITI PFA.
 * @param {Object} projectVar instance of getProjectVariable()
 * @param {Object} daerahList pass daerah object in a list
 * @param {Object} targetTable instance of Table (Google Docs)
 */
function fetchPfa(projectVar, daerahList, targetTable) {
  let sourceRange = projectVar.sheetPfa.getRange('B14:M25');
  let sourceValues = sourceRange.getValues();
  daerahList.forEach(daerah => {
    let values = sourceValues[daerah.position];
    // Think of the `appendTableRow()` as <tr> in HTML where you have to set the cell using <td> which is `appendTableCell()`
    let newTableRow = targetTable.appendTableRow();
    values.forEach(item => {
      if (typeof item == 'number') {
        newTableRow.appendTableCell(item.toFixed());
      } else {
        newTableRow.appendTableCell(item);
      }
    });
  });
  let sumValues = sourceValues[11];
  let sumTableRow = targetTable.appendTableRow();
  sumValues.forEach(item => {
    if (typeof item == 'number') {
      sumTableRow.appendTableCell(item.toFixed());
    } else {
      sumTableRow.appendTableCell(item);
    }
  });
  targetTable.setBorderWidth(0.75);
}
