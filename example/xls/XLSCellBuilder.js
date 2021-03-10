const xlsUtils = require('../../index').XLSUtils;

const HEADERS = ['Designator', 'Hierarchy', 'Field 1', 'Field 2', 'Field 3'];

const FIELDS = [
  { n: 'designator', t: 'string' },
  { n: 'heirarchy', t: 'string' },
  { n: 'field1', t: 'string' },
  { n: 'field2', t: 'string' },
  { n: 'field3', t: 'string' }
];

module.exports = {
  build: function(wb, data, styles) {
    const worksheet = wb.addWorksheet('Field Report');
    buildHeaders(worksheet, data, styles);
    buildBody(worksheet, data, styles);
  }
};

const buildHeaders = function(sheet, data, styles) {
  HEADERS.forEach(function(h, i) {
    xlsUtils.buildCell(sheet, h, 'string', xlsUtils.incrementFromCell('A1', 0, i), null, false, [styles.headerStyle, styles.boldStyle]);
  });
};

const buildBody = function(sheet, data, styles) {
  data.forEach(function(t, rowIncrement) {
    for(let columnIncrement = 0; columnIncrement < FIELDS.length; columnIncrement++) {
      let fieldName = FIELDS[columnIncrement].n;
      let fieldType = FIELDS[columnIncrement].t;
      let fieldValue = (fieldType === 'string') ? t[fieldName] || '' : parseFloat(t[fieldName]);
      let style = null;
      xlsUtils.buildCell(sheet, fieldValue, fieldType, xlsUtils.incrementFromCell('A2', rowIncrement, columnIncrement), null, false, style);
    }
  });
};
