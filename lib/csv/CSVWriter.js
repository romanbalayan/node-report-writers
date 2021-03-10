const fs = require('fs');
const path = require('path');

function CSVWriter(csvWriterParams, settlementData) {
  this.settlementData = settlementData;
  this.writerHelper = csvWriterParams.writerHelper;
  this.filepath = csvWriterParams.filepath;
  this.filename = csvWriterParams.filename;
}

CSVWriter.prototype.write = function(callback) {
  const filename = path.join(this.filepath, this.filename);

  const writer = this.writerHelper;
  const settlementData = this.settlementData;
  const stream = fs.createWriteStream(filename);

  stream.once('open', function() {
    writer.write(settlementData, stream);
    stream.write('');
    stream.end();
    stream.on('finish', function() {
      callback(null, {filename: filename});
    });
  });
};

module.exports = CSVWriter;
