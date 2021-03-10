/**
 *
 * @param str
 * @returns {{row: Number, col: Number}}
 */
function getExcelRowCol(str) {
  const numeric = str.split(/\D/).filter(function (el) {
    return el !== '';
  })[0];
  const alpha = str.split(/\d/).filter(function (el) {
    return el !== '';
  })[0];
  const row = parseInt(numeric, 10);
  const col = alpha.toUpperCase().split('').reduce(function (a, b, index, arr) {
    return a + (b.charCodeAt(0) - 64) * Math.pow(26, arr.length - index - 1);
  }, 0);
  return {
    row: row,
    col: col
  };
}

/**
 *
 * @param column
 * @param row
 * @returns {string}
 */
function getCellFromColumnAndRow(column, row) {
  const aCharCode = 65;
  let remaining = column;
  let columnString = '';
  while (remaining > 0) {
    let mod = (remaining - 1) % 26;
    columnString = String.fromCharCode(aCharCode + mod) + columnString;
    remaining = (remaining - 1 - mod) / 26;
  }
  return columnString + row;
}

/**
 *
 * @param startCell
 * @param rowIncrement
 * @param colIncrement
 * @returns {string}
 */
function incrementFromCell(startCell, rowIncrement, colIncrement) {
  rowIncrement = rowIncrement || 0;
  colIncrement = colIncrement || 0;
  const start = getExcelRowCol(startCell);
  return getCellFromColumnAndRow(start.col + colIncrement, start.row + rowIncrement);
}

/**
 *
 * @param ws
 * @param value
 * @param type
 * @param startCell
 * @param endCell
 * @param isMerged
 * @param style
 * @param width
 * @param height
 */
function buildCell(ws, value, type, startCell, endCell, isMerged, style, width, height) {
  const cell = _buildCell(value, type, startCell, endCell, isMerged, style, width, height);
  _writeToCell(ws, cell, cell.style);
}

/**
 *
 * @param ws
 * @param row
 * @param maxColumns
 */
function addFillerRow(ws, row, maxColumns) {
  const startCell = getCellFromColumnAndRow(1, row);
  const endCell = getCellFromColumnAndRow(maxColumns, row);
  return buildCell(ws, '', 'string', startCell, endCell, true, null, null, 4);
}


module.exports = {
  incrementFromCell: incrementFromCell,
  getCellFromColumnAndRow: getCellFromColumnAndRow,
  buildCell: buildCell,
  addFillerRow: addFillerRow
};


/**
 *
 * @param value
 * @param type
 * @param startCell
 * @param endCell
 * @param isMerged
 * @param style
 * @param width
 * @param height
 * @returns {{value: string, type: *, startRow: Number, startColumn: Number, endRow: Number, endColumn: Number, isMerged: (*|boolean), style: *, width: *, height: *}}
 * @private
 */
function _buildCell(value, type, startCell, endCell, isMerged, style, width, height) {
  endCell = endCell || startCell;
  const start = getExcelRowCol(startCell);
  const end = getExcelRowCol(endCell);
  return {
    value: value == null ? '' : value,
    type: type,
    startRow: start.row,
    startColumn: start.col,
    endRow: end.row,
    endColumn: end.col,
    isMerged: isMerged || false,
    style: style,
    width: width,
    height: height
  };
}

/**
 *
 * @param sheet
 * @param cellDetails
 * @param style
 * @private
 */
function _writeToCell(sheet, cellDetails, style) {
  if(cellDetails.width) {
    sheet.getColumn(cellDetails.startColumn).width = cellDetails.width;
  }

  if(cellDetails.height) {
    sheet.getRow(cellDetails.startRow).height = cellDetails.height;
  }

  const startCell = getCellFromColumnAndRow(cellDetails.startColumn, cellDetails.startRow);

  if(cellDetails.isMerged) {
    const endCell = getCellFromColumnAndRow(cellDetails.endColumn, cellDetails.endRow);
    sheet.mergeCells(startCell + ':' + endCell);
  }

  sheet.getCell(startCell).value = cellDetails.value;

  if(style) {
    if(style.constructor === Array) {
      style.forEach(function(s) {
        const styles = Object.keys(s);
        styles.forEach(function(_s) {
          sheet.getCell(startCell)[_s] = s[_s];
        });
      });
    } else {
      const styles = Object.keys(style);
      styles.forEach(function(s) {
        sheet.getCell(startCell)[s] = style[s];
      });
    }
  }
}
