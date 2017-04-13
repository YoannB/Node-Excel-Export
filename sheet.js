/**
 * version 0.5.2
 *
 * 2017-04-13 : YoannB
 */
var sheetFront = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    + 'mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
    + ' <sheetViews><sheetView tabSelected="1" workbookViewId="0" /></sheetViews>'
    + ' <sheetFormatPr baseColWidth="10" defaultColWidth="10" defaultRowHeight="15" x14ac:dyDescent="0"/>';
var sheetBack = '<pageMargins left="0.75" right="0.75" top="0.75" bottom="0.5" header="0.5" footer="0.75" />'
    + '</worksheet>';

var fs = require('fs');

function Sheet(config, xlsx, shareStrings, convertedShareStrings) {
    this.config = config;
    this.xlsx = xlsx;
    this.shareStrings = shareStrings;
    this.convertedShareStrings = convertedShareStrings;
}

Sheet.prototype.generate = function () {
    // fix for different cols width with multiple export
    var sheetFrontClone = sheetFront.slice(0);

    var config = this.config, xlsx = this.xlsx;
    var cols = config.cols,
        data = config.rows,
        colsLength = cols.length,
        rows = "",
        row = "",
        colsWidth = "",
        styleIndex,
        self = this,
        k;
    config.fileName = 'xl/worksheets/' + (config.name || "sheet").replace(/[*?\]\[\/\/]/g, '') + '.xml';
    if (config.stylesXmlFile) {
        var path = config.stylesXmlFile;
        var styles = fs.readFileSync(path, 'utf8');
        if (styles) {
            xlsx.file("xl/styles.xml", styles);
        }
    }

    //first row for column caption
    row = '<row r="1" spans="1:' + colsLength + '" ht="20">';
    var colStyleIndex;
    for (k = 0; k < colsLength; k++) {
        colStyleIndex = cols[k].captionStyleIndex || 0;
        row += addStringCell(self, getColumnLetter(k + 1) + 1, cols[k].caption, colStyleIndex);
        if (cols[k].width) {
            colsWidth += '<col customWidth="1" width="' + cols[k].width + '" max="' + (k + 1) + '" min="' + (k + 1) + '"/>';
        }
    }
    row += '</row>';
    rows += row;

    //fill in data
    var i, j, r, cellData, currRow, cellType, dataLength = data.length;

    for (i = 0; i < dataLength; i++) {
        r = data[i];
        currRow = i + 2;
        row = '<row r="' + currRow + '" spans="1:' + colsLength + '" ht="18">';
        for (j = 0; j < colsLength; j++) {
            styleIndex = null;
            cellData = r[j];
            cellType = cols[j].type;
            if (typeof cols[j].beforeCellWrite === 'function') {
                var e = {
                    rowNum: currRow,
                    styleIndex: null,
                    cellType: cellType
                };
                cellData = cols[j].beforeCellWrite(r, cellData, e);
                styleIndex = e.styleIndex || styleIndex;
                cellType = e.cellType;
                delete e;
            }
            switch (cellType) {
                case 'number':
                    row += addNumberCell(getColumnLetter(j + 1) + currRow, cellData, styleIndex);
                    break;
                case 'date':
                    row += addDateCell(getColumnLetter(j + 1) + currRow, cellData, styleIndex);
                    break;
                case 'bool':
                    row += addBoolCell(getColumnLetter(j + 1) + currRow, cellData, styleIndex);
                    break;
                default:
                    row += addStringCell(self, getColumnLetter(j + 1) + currRow, cellData, styleIndex);
            }
        }
        row += '</row>';
        rows += row;
    }
    if (colsWidth !== "") {
        sheetFrontClone += '<cols>' + colsWidth + '</cols>';
    }
    xlsx.file(config.fileName, sheetFrontClone + '<sheetData>' + rows + '</sheetData>' + sheetBack);
};

module.exports = Sheet;

var startTag = function (obj, tagName, closed) {
    var result = "<" + tagName, p;
    for (p in obj) {
        result += " " + p + "=" + obj[p];
    }
    if (!closed)
        result += ">";
    else
        result += "/>";
    return result;
};

var endTag = function (tagName) {
    return "</" + tagName + ">";
};

var addNumberCell = function (cellRef, value, styleIndex) {
    styleIndex = styleIndex || 0;
    if (value === null) {
        return "";
    }  else {
        return '<c r="' + cellRef + '"' + (styleIndex > 0 ? ' s="' + styleIndex + '"' : '') + ' t="n"><v>' + value + '</v></c>';
    }
};

var addDateCell = function (cellRef, value, styleIndex) {
    styleIndex = styleIndex || 1;
    if (value === null) {
        return "";
    } else {
        return '<c r="' + cellRef + '"' + (styleIndex > 0 ? ' s="' + styleIndex + '"' : '') + ' t="n"><v>' + value + '</v></c>';
    }
};

var addBoolCell = function (cellRef, value, styleIndex) {
    styleIndex = styleIndex || 0;
    if (value === null) {
        return "";
    }
    if (value) {
        value = 1;
    } else {
        value = 0;
    }
    return '<c r="' + cellRef + '"' + (styleIndex > 0 ? ' s="' + styleIndex + '"' : '') + ' t="b"><v>' + value + '</v></c>';
};

var addStringCell = function (sheet, cellRef, value, styleIndex) {
    styleIndex = styleIndex || 0;
    if (value === null)
        return "";
    if (typeof value === 'string') {
        value = value.replace(/&/g, "&amp;").replace(/'/g, "&apos;").replace(/>/g, "&gt;").replace(/</g, "&lt;");
    }
    var i = sheet.shareStrings[value];
    if (!i) {
        i = Object.keys(sheet.shareStrings).length;
        sheet.shareStrings[value] = i;
        sheet.convertedShareStrings += "<si><t>" + value + "</t></si>";
    }
    return '<c r="' + cellRef + '"' + (styleIndex > 0 ? ' s="' + styleIndex + '"' : '') + ' t="s"><v>' + i + '</v></c>';
};

var getColumnLetter = function (col) {
    if (col <= 0)
        throw "col must be more than 0";
    var array = [];
    while (col > 0) {
        var remainder = col % 26;
        col /= 26;
        col = Math.floor(col);
        if (remainder === 0) {
            remainder = 26;
            col--;
        }
        array.push(64 + remainder);
    }
    return String.fromCharCode.apply(null, array.reverse());
};

