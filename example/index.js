const async = require('async');
const nodeReportWriters = require('../index');

const XLSWriter = nodeReportWriters.XLSWriter;
const CSVWriter = nodeReportWriters.CSVWriter;

const xlsStyles = require('./xls/XLSStyles');
const XLSCellBuilder = require('./xls/XLSCellBuilder');

const CSVWriterHelper = require('./csv/CSVWriterHelper');

const data = require('./data.json');
const FILE_PATH = __dirname;

async.auto({
  createCSV: function (asyncCb) {
    console.log(new Date(), 'Creating CSV Report.');

    const params = {
      filename: 'example.csv',
      filepath: FILE_PATH,
      writerHelper: CSVWriterHelper
    };

    const csvWriter = new CSVWriter(params, data);

    csvWriter.write(function (err, result) {
      if (err) {
        return asyncCb(err);
      }
      asyncCb(null, result);
    });
  },

  createXLS: function (asyncCb) {
    console.log(new Date(), 'Creating XLS Report.');

    const filename = 'example.xlsx';
    const params = {
      filename: filename,
      filepath: FILE_PATH,
      cellBuilder: XLSCellBuilder,
      styles: xlsStyles
    };

    const xlsWriter = new XLSWriter(params, data);
    xlsWriter.write(function (err, result) {
      if (err) {
        return asyncCb(err);
      }
      asyncCb(null, result);
    });
  }
}, function done(err, results) {
  console.log(err);
  console.log(results);
});
