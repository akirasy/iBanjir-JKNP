function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('iBanjir JKNP')
  .addItem('🕸 Set Google permission', 'aquireGooglePermission')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🕸 Formatting')
    .addItem('⚪ To upperCase', 'toUpperCase')
    .addItem('⚪ To oneLine', 'toOneLine')
    .addItem('⚪ Clean IC', 'cleanIc'))
  .addSeparator()
  .addItem('Generate laporan banjir', 'generateNow')
  .addToUi();
}

/**
 * Collect and generate laporan harian banjir.
 */
function generateNow() {
  let projectVar      = getProjectVariable();
  let templateLaporan = DriveApp.getFileById(projectVar.template_laporan_id);
  let outputFolder    = DriveApp.getFolderById(projectVar.output_folder_id);
  let daerahBanjir    = new Array();
  projectVar.listDaerah.forEach(item => {
    if (projectVar.dictDaerah[item]['isBanjir'] == true) {
      daerahBanjir.push(projectVar.dictDaerah[item]);
    }
  })
  generateLaporan(projectVar, daerahBanjir, templateLaporan, outputFolder);
}

/**
 * Generate laporan harian banjir.
 * @param {Object} projectVar instance of getProjectVariable()
 * @param {Object} daerahList pass daerah object in a list
 * @param {Object} templateLaporan pass a DriveApp-File object
 * @param {Object} outputFolder pass a DriveApp-Folder object
 */
function generateLaporan(projectVar, daerahList, templateLaporan, outputFolder) {
  let today = new Date();
  let filename = 'SITUASI KEJADIAN BANJIR NEGERI PAHANG ' + today.toLocaleString();
  let newFile = templateLaporan.makeCopy(filename, outputFolder);
  let newFileDoc = DocumentApp.openById(newFile.getId());

  let allTemplateTables = newFileDoc.getBody().getTables();
  Logger.log('Generate table for Denggi');
  let tableDengue = allTemplateTables[2];
  fetchDengueData(projectVar, daerahList, tableDengue);
  Logger.log('Generate table for Lalat');
  let tableLalat = allTemplateTables[3];
  fetchLalatData(projectVar, daerahList, tableLalat);
  Logger.log('Generate table for LipasTikus');
  let tableLiTi = allTemplateTables[4];
  fetchLalatData(projectVar, daerahList, tableLiTi);
  Logger.log('Generate table for Pendidikan Kesihatan');
  let tableHeccOne = allTemplateTables[5];
  fetchHeccOne(projectVar, daerahList, tableHeccOne);
  let tableHeccTwo = allTemplateTables[6];
  fetchHeccTwo(projectVar, daerahList, tableHeccTwo);
  Logger.log('Generate table for Psychological First Aid');
  let tablePfa = allTemplateTables[7];
  fetchPfa(projectVar, daerahList, tablePfa);
}


















