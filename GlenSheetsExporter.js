/**
 * GlenSheetsExporter
 * Louille Glen Benatiro
 * June 2024
 * glenbenatiro@gmail.com
 */

// =============================================================================

// ESLint globals
/* global DriveApp */
/* global ScriptApp */
/* global UrlFetchApp */
/* global SpreadsheetApp */

// =============================================================================

/**
 * parameter references
 * 1. https://stackoverflow.com/questions/46088042/margins-parameters-for-google-spreadsheet-export-as-pdf\
 * 2. https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
 * 3. https://gist.github.com/andrewroberts/c37d45619d5661cab078be2a3f2fd2bb
 */

// =============================================================================

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

// eslint-disable-next-line no-var
var ExportRangeType = {
  SHEET: '0',
  WORKBOOK: '1',
  RANGE: '2',
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
  PRINT_NOTES: 'printnotes',
  SHEET_ID: 'gid',
  IR: 'ir',
  IC: 'ic',
  R1: 'r1',
  C1: 'c1',
  R2: 'r2',
  C2: 'c2',
};

// eslint-disable-next-line no-var
var DEFAULT_EXPORT_SETTINGS = {
  [EXPORT_SETTINGS.FORMAT]: Format.PDF,
  [EXPORT_SETTINGS.SIZE]: Size.A4,
  [EXPORT_SETTINGS.ORIENTATION]: Orientation.PORTRAIT,
  [EXPORT_SETTINGS.SCALE]: Scale.FIT_TO_PAGE,
  [EXPORT_SETTINGS.TOP_MARGIN]: 0.5,
  [EXPORT_SETTINGS.BOTTOM_MARGIN]: 0.5,
  [EXPORT_SETTINGS.LEFT_MARGIN]: 0.5,
  [EXPORT_SETTINGS.RIGHT_MARGIN]: 0.5,
  [EXPORT_SETTINGS.REPEAT_ROW_HEADERS]: false,
  [EXPORT_SETTINGS.GRIDLINES]: false,
  [EXPORT_SETTINGS.PRINT_TITLE]: false,
  [EXPORT_SETTINGS.PRINT_NOTES]: false,
};

const DEFAULT_RUNTIME_EXPORT_SETTINGS = {
  exportRange: {
    type: ExportRangeType.WORKBOOK,
    sheetName: null,
    sheetIndex: null,
    sheetRange: null,
  },
};

const GLENSHEETSTOPDF_DEFAULT_EXPORT_SETTINGS = {
  actual: DEFAULT_EXPORT_SETTINGS,
  runtime: DEFAULT_RUNTIME_EXPORT_SETTINGS,
};

// =============================================================================

function exportSpreadsheet_(spreadsheet, exportSettings) {
  const spreadsheetID = spreadsheet.getId();
  const baseURL = `https://docs.google.com/spreadsheets/d/${spreadsheetID}/export?`;
  const queryParams = Object.entries(exportSettings)
    .reduce((accumulator, [key, value]) => {
      accumulator.push(`${key}=${value}`);
      return accumulator;
    }, [])
    .join('&');
  const exportURL = `${baseURL}${queryParams}`;
  const response = UrlFetchApp.fetch(exportURL, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
  });
  const exportFile = DriveApp.createFile(response.getBlob()).setName(
    spreadsheet.getName(),
  );

  return exportFile;
}


function sheetColumnLettersToNumber(column) {
  let number = 0;

  for (let i = 0; i < column.length; i++) {
      number = number * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }

  return number;
}

function cellToRC(cell) {
  const match = cell.match(/([A-Z]+)(\d+)/);

  return {
      column: sheetColumnLettersToNumber(match[1]),
      row: parseInt(match[2], 10),
  };
}

function sheetA1RangeToRC12_(a1Notation) {
  const rangeParts = a1Notation.split(':');
  const startCell = cellToRC(rangeParts[0]);
  
  let endCell;
  if (rangeParts.length > 1) {
      endCell = cellToRC(rangeParts[1]);
  } else {
      endCell = startCell;
  }

  return {
      R1: startCell.row,  // First row
      C1: startCell.column,  // First column
      R2: endCell.row,    // Last row
      C2: endCell.column  // Last column
  };
}

// =============================================================================

class GlenSheetsExplorer {
  constructor() {
    this.exportSettings_ = GLENSHEETSTOPDF_DEFAULT_EXPORT_SETTINGS;
  }

  setSize(size) {
    this.exportSettings_.actual[EXPORT_SETTINGS.SIZE] = size;
    return this;
  }

  setMargins(
    top = this.exportSettings_.actual[EXPORT_SETTINGS.TOP_MARGIN],
    bottom = this.exportSettings_.actual[EXPORT_SETTINGS.BOTTOM_MARGIN],
    left = this.exportSettings_.actual[EXPORT_SETTINGS.LEFT_MARGIN],
    right = this.exportSettings_.actual[EXPORT_SETTINGS.RIGHT_MARGIN],
  ) {
    this.exportSettings_.actual[EXPORT_SETTINGS.TOP_MARGIN] = top;
    this.exportSettings_.actual[EXPORT_SETTINGS.BOTTOM_MARGIN] = bottom;
    this.exportSettings_.actual[EXPORT_SETTINGS.LEFT_MARGIN] = left;
    this.exportSettings_.actual[EXPORT_SETTINGS.RIGHT_MARGIN] = right;

    return this;
  }

  setScale(scale) {
    this.exportSettings_.actual[EXPORT_SETTINGS.SCALE] = scale;
    return this;
  }

  setExportRange(exportRangeType, param1, param2) {
    switch (exportRangeType) {
      case ExportRangeType.WORKBOOK:
        break;

      case ExportRangeType.SHEET:
      case ExportRangeType.RANGE: {
        switch (typeof param1) {
          case 'string':
            this.exportSettings_.runtime.exportRange.sheetName = param1;
            break;

          case 'number':
            this.exportSettings_.runtime.exportRange.sheetIndex = param1;
            break;

          default:
            throw new Error(
              `Invalid type for argument param1: ${typeof param1}. Expecting string or number.`,
            );
        }

        if (exportRangeType === ExportRangeType.RANGE) {
          this.exportSettings_.runtime.exportRange.sheetRange = param2;
        }

        break;
      }

      default: {
        throw new Error(
          `Invalid argument for exportRangeType: ${exportRangeType}`,
        );
      }
    }

    this.exportSettings_.runtime.exportRange.type = exportRangeType;
    return this;
  }

  preExport_(spreadsheet) {
    const exportSettings = this.exportSettings_.actual;

    delete exportSettings[EXPORT_SETTINGS.SHEET_ID];
    delete exportSettings[EXPORT_SETTINGS.IR];
    delete exportSettings[EXPORT_SETTINGS.IC];
    delete exportSettings[EXPORT_SETTINGS.R1];
    delete exportSettings[EXPORT_SETTINGS.C1];
    delete exportSettings[EXPORT_SETTINGS.R2];
    delete exportSettings[EXPORT_SETTINGS.C2];

    const exportRangeType = this.exportSettings_.runtime.exportRange.type;

    switch (exportRangeType) {
      case ExportRangeType.WORKBOOK:
        break;

      case ExportRangeType.SHEET:
      case ExportRangeType.RANGE: {
        const targetSheetName =
          this.exportSettings_.runtime.exportRange.sheetName;
        const sheet = spreadsheet
          .getSheets()
          .find((curr) => curr.getName() === targetSheetName);

        if (!sheet) {
          throw new Error(
            `ExportRangeType.SHEET sheet name ${targetSheetName} not found in spreadsheet.`,
          );
        } else {
          exportSettings[EXPORT_SETTINGS.SHEET_ID] = sheet.getSheetId();
        }

        if (exportRangeType === ExportRangeType.RANGE) {
          exportSettings[EXPORT_SETTINGS.IR] = false;
          exportSettings[EXPORT_SETTINGS.IC] = false;

          const rc12 = sheetA1RangeToRC12_(this.exportSettings_.runtime.exportRange.sheetRange);

          exportSettings[EXPORT_SETTINGS.R1] = rc12.R1 - 1;
          exportSettings[EXPORT_SETTINGS.C1] = rc12.C1 - 1;
          exportSettings[EXPORT_SETTINGS.R2] = rc12.R2;
          exportSettings[EXPORT_SETTINGS.C2] = rc12.C2;
        }

        break;
      }

      default:
        throw new Error();
    }

    return exportSpreadsheet_(spreadsheet, exportSettings);
  }

  exportBySpreadsheet(spreadsheet) {
    return this.preExport_(spreadsheet);
  }

  exportByURL(url) {
    return this.exportBySpreadsheet(SpreadsheetApp.openByUrl(url));
  }
}

// =============================================================================

function createInstance() {
  return new GlenSheetsExplorer();
}

// EOF
