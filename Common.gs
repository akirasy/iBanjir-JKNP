/**
 * Instantiate global project variable to save execution time.
 */
function getProjectVariable() {
  let output = new Object();

  // Instantiate GSheet
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheetAppscript = activeSpreadsheet.getSheetByName('appscript.gs');
  let sheetDenggi = activeSpreadsheet.getSheetByName('BRG 8 DENGGI');
  let sheetLalat = activeSpreadsheet.getSheetByName('BRG 9 LALAT');
  let sheetLiTi = activeSpreadsheet.getSheetByName('BRG 10 LITI');
  let sheetHeccOne = activeSpreadsheet.getSheetByName('BRG 5.1 HECC');
  let sheetHeccTwo = activeSpreadsheet.getSheetByName('BRG 5.2 HECC');
  let sheetPfa = activeSpreadsheet.getSheetByName('BRG 17 AKTIVITI PFA');
  output['activeSpreadsheet']    = activeSpreadsheet;
  output['sheetAppscript']       = sheetAppscript;
  output['sheetDenggi']          = sheetDenggi;
  output['sheetLalat']           = sheetLalat;
  output['sheetLiTi']            = sheetLiTi;
  output['sheetHeccOne']         = sheetHeccOne;
  output['sheetHeccTwo']         = sheetHeccTwo;
  output['sheetPfa']             = sheetPfa;

  // Obtain user-defined values in sheet `appscript.gs`
  let varKeyValue  = sheetAppscript.getRange(2, 1, sheetAppscript.getLastRow(), 2).getValues();
  varKeyValue.forEach(item => {
    if (item[0]) {
      output[item[0]] = item[1];
    }
  });

  // Map dictionary of daerah into callable info
  output['listDaerah']        = ['bentong', 'bera', 'cameron', 'jerantut', 'lipis', 'kuantan', 'maran', 'pekan', 'raub', 'rompin', 'temerloh']
  output['dictDaerah']        = {
    'bentong' : {'position':0, 'name':'Bentong', 'isBanjir':output.is_banjir_bentong},
    'bera'    : {'position':1, 'name':'Bera', 'isBanjir':output.is_banjir_bera},
    'cameron' : {'position':2, 'name':'Cameron Highland', 'isBanjir':output.is_banjir_cameron},
    'jerantut': {'position':3, 'name':'Jerantut', 'isBanjir':output.is_banjir_jerantut},
    'lipis'   : {'position':4, 'name':'Lipis', 'isBanjir':output.is_banjir_lipis},
    'kuantan' : {'position':5, 'name':'Kuantan', 'isBanjir':output.is_banjir_kuantan},
    'maran'   : {'position':6, 'name':'Maran', 'isBanjir':output.is_banjir_maran},
    'pekan'   : {'position':7, 'name':'Pekan', 'isBanjir':output.is_banjir_pekan},
    'raub'    : {'position':8, 'name':'Raub', 'isBanjir':output.is_banjir_raub},
    'rompin'  : {'position':9, 'name':'Rompin', 'isBanjir':output.is_banjir_rompin},
    'temerloh': {'position':10, 'name':'Temerloh', 'isBanjir':output.is_banjir_temerloh},
    };

  return output
}

/**
 * Check if user has approve script execution.
 */
function aquireGooglePermission() {
  SpreadsheetApp.getUi().alert(
    'Success',
    'If you can see this. You already have permission to use this app.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Change lowercase to uppercase.
 */
function toUpperCase() {
  let activeSheet = SpreadsheetApp.getActiveSheet();
  let selected_range = activeSheet.getActiveRange();
  let data_list = selected_range.getValues();
  for (let i = 0; i < data_list.length; i++) {
    for (let j = 0; j < data_list[i].length; j++) {
      if (!(data_list[i][j] instanceof Date)) {
        data_list[i][j] = data_list[i][j].toString().toUpperCase();
      }
    }
  }
  selected_range.setValues(data_list);
}

/**
 * Convert newline value to oneline only.
 */
function toOneLine() {
  let activeSheet = SpreadsheetApp.getActiveSheet();
  let selected_range = activeSheet.getActiveRange();
  let data_list = selected_range.getValues();
  for (let i = 0; i < data_list.length; i++) {
    for (let j = 0; j < data_list[i].length; j++) {
      if (!(data_list[i][j] instanceof Date)) {
        data_list[i][j] = data_list[i][j].toString().replace(/\n/g, '  ');
      }
    }
  }
  selected_range.setValues(data_list).trimWhitespace();
}

/**
 * Removes dashes, star, spaces and apostrophy.
 */
function cleanIc() {
  let activeSheet = SpreadsheetApp.getActiveSheet();
  let selected_range = activeSheet.getActiveRange();
  let data_list = selected_range.getValues();
  for (let i = 0; i < data_list.length; i++) {
    for (let j = 0; j < data_list[i].length; j++) {
      if (!(data_list[i][j] instanceof Date)) {
        data_list[i][j] = data_list[i][j].toString().replace(/[-|\'|\*|\s]/g,'');
      }
    }
  }
  selected_range.setValues(data_list);
}
