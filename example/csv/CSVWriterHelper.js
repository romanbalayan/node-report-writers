module.exports = {
  write: function(data, writeStream) {
    headerWriter(data, writeStream);
    bodyWriter(data, writeStream);
  }
};

function headerWriter(data, writeStream) {
  const headers = Object.keys(data[0]).map(function(h) {
    return h.toUpperCase();
  });
  writeStream.write(headers.toString() + '\n');
}

function bodyWriter(data, writeStream) {
  function csvify(data, fields) {
    return fields.reduce(function(accumulator, f) {
      accumulator += (data[f] == null) ? '' : data[f];
      if (fields.indexOf(f) + 1 !== fields.length) {
        accumulator += ',';
      }
      return accumulator;
    }, '');
  }

  const fields = Object.keys(data[0]);

  data.forEach(function(d) {
    const csvString = csvify(d, fields);
    writeStream.write(csvString + '\n');
  });
}
