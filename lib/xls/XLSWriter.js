const path = require('path');
const excel = require('exceljs');

function XLSWriter(xlsWriterParams, sourceData) {
  this.sourceData = sourceData;
  this.cellBuilder = xlsWriterParams.cellBuilder;
  this.filename = xlsWriterParams.filename;
  this.filepath = xlsWriterParams.filepath;
  this.styles = xlsWriterParams.styles;

  const _filename = path.join(this.filepath, this.filename);
  const options = {
    filename: _filename,
    useStyles: true,
    useSharedStrings: true
  };

  this.wb = new excel.stream.xlsx.WorkbookWriter(options);
}


XLSWriter.prototype.write = function(callback) {
  const wb = this.wb;
  const styles = this.styles;
  const sourceData = this.sourceData;
  const cellBuilder = this.cellBuilder;
  const _filename = path.join(this.filepath, this.filename);

  cellBuilder.build(wb, sourceData, styles);

  this.wb
  .commit()
  .then(function() {
    callback(null, {filename: _filename});
  }, function(err) {
    callback(err);
  });
};

module.exports = XLSWriter;
