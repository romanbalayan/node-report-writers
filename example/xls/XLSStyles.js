const blackBorder = {
  style: 'thin',
  color: {argb: 'FF000000'}
};

const borderStyle = {
  left: blackBorder,
  right: blackBorder,
  top: blackBorder,
  bottom: blackBorder
};

module.exports = {
  headerStyle: {
    alignment: {
      horizontal: 'center'
    },
    border: borderStyle
  },
  bodyStyle: {
    border: borderStyle
  },
  boldStyle: {
    font: {
      bold: true
    }
  }
};
