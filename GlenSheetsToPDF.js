/**
 * GlenSheetsToPDF
 * Louille Glen Benatiro
 * June 2024
 * glenbenatiro@gmail.com
 */

// ESLint globals
/* global DriveApp */
/* global ScriptApp */
/* global UrlFetchApp */
/* global SpreadsheetApp */

// eslint-disable-next-line no-var
var Format = {
  XLSX: 'xlsx',
  ODS: 'ods',
  ZIP: 'zip',
  CSV: 'csv',
  TSV: 'tsv',
  PDF: 'pdf',
};

// eslint-disable-next-line no-var
var Size = {
  LETTER: 0,
  TABLOID: 1,
  LEGAL: 2,
  STATEMENT: 3,
  EXECUTIVE: 4,
  FOLIO: 5,
  A3: 6,
  A4: 7,
  A5: 8,
  B4: 9,
  B5: 10,
};

// eslint-disable-next-line no-var
var Orientation = {
  PORTRAIT: 'true',
  LANDSCAPE: 'false',
};

// eslint-disable-next-line no-var
var Scale = {
  NORMAL: '1',
  FIT_TO_WIDTH: '2',
  FIT_TO_HEIGHT: '3',
  FIT_TO_PAGE: '4',
};

// =============================================================================

// eslint-disable-next-line no-var
var EXPORT_SETTINGS = {
  FORMAT: 'format',
  SIZE: 'size',
  REPEAT_ROW_HEADERS: 'fzr',
  ORIENTATION: 'portrait',
  GRIDLINES: 'gridlines',
  PRINT_TITLE: 'printtitle',
  SCALE: 'scale',
  TOP_MARGIN: 'top_margin',
  BOTTOM_MARGIN: 'bottom_margin',
  LEFT_MARGIN: 'left_margin',
  RIGHT_MARGIN: 'right_margin',
};

// eslint-disable-next-line no-var
var DEFAULT_EXPORT_SETTINGS = {
  [EXPORT_SETTINGS.FORMAT]: Format.PDF,
  [EXPORT_SETTINGS.SIZE]: Size.A4,
  [EXPORT_SETTINGS.REPEAT_ROW_HEADERS]: false,
  [EXPORT_SETTINGS.ORIENTATION]: Orientation.PORTRAIT,
  [EXPORT_SETTINGS.SCALE]: Scale.FIT_TO_PAGE,
  [EXPORT_SETTINGS.TOP_MARGIN]: 0.5,
  [EXPORT_SETTINGS.BOTTOM_MARGIN]: 0.5,
  [EXPORT_SETTINGS.LEFT_MARGIN]: 0.5,
  [EXPORT_SETTINGS.RIGHT_MARGIN]: 0.5,
  [EXPORT_SETTINGS.GRIDLINES]: false,
  [EXPORT_SETTINGS.PRINT_TITLE]: false,
};

// =============================================================================

function convertSheetsToPDF_(spreadsheet, exportSettings, destFolder) {
  const spreadsheetID = spreadsheet.getId();
  const baseURL = `https://docs.google.com/spreadsheets/d/${spreadsheetID}/export?`;
  const settings = exportSettings ?? DEFAULT_EXPORT_SETTINGS;
  const queryParams = Object.entries(settings)
    .reduce((accumulator, [key, value]) => {
      accumulator.push(`${key}=${value}`);
      return accumulator;
    }, [])
    .join('&');
  const exportURL = `${baseURL}${queryParams}`;
  const response = UrlFetchApp.fetch(exportURL, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
  });
  const folder =
    destFolder ?? DriveApp.getFileById(spreadsheet.getId()).getParents().next();
  const pdf = DriveApp.createFile(response.getBlob())
    .moveTo(folder)
    .setName(spreadsheet.getName());

  return pdf;
}

// =============================================================================

class GlenSheetToPDF {
  constructor() {
    // storage
    this.destinationFolder_ = null;

    // conversion settings
    this.exportSettings_ = DEFAULT_EXPORT_SETTINGS;

    // state
    this.deleteSpreadsheetAfterConversion_ = true;
  }

  setSize(size) {
    this.exportSettings_[EXPORT_SETTINGS.SIZE] = size;
    return this;
  }

  setFormat(format) {
    this.exportSettings_[EXPORT_SETTINGS.FORMAT] = format;
    return this;
  }

  setMargins(
    top = this.exportSettings_[EXPORT_SETTINGS.TOP_MARGIN],
    bottom = this.exportSettings_[EXPORT_SETTINGS.BOTTOM_MARGIN],
    left = this.exportSettings_[EXPORT_SETTINGS.LEFT_MARGIN],
    right = this.exportSettings_[EXPORT_SETTINGS.RIGHT_MARGIN],
  ) {
    this.exportSettings_[EXPORT_SETTINGS.TOP_MARGIN] = top;
    this.exportSettings_[EXPORT_SETTINGS.BOTTOM_MARGIN] = bottom;
    this.exportSettings_[EXPORT_SETTINGS.LEFT_MARGIN] = left;
    this.exportSettings_[EXPORT_SETTINGS.RIGHT_MARGIN] = right;

    return this;
  }

  setScale(scale) {
    this.exportSettings_[EXPORT_SETTINGS.SCALE] = scale;
    return this;
  }

  convert(spreadsheet) {
    return convertSheetsToPDF_(
      spreadsheet,
      this.exportSettings_,
      this.destinationFolder_,
    );
  }

  convertByURL(spreadsheetURL) {
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetURL);

    return this.convert(spreadsheet);
  }
}

// =============================================================================

function createInstance() {
  return new GlenSheetToPDF();
}

function convert(spreadsheet, exportSettings, destFolder) {
  return convertSheetsToPDF_(spreadsheet, exportSettings, destFolder);
}

function convertByURL(spreadsheetURL, exportSettings, destFolder) {
  return convert(SpreadsheetApp.openByUrl(spreadsheetURL));
}

// EOF
