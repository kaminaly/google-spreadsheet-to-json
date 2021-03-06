'use strict';

var GoogleSpreadsheet = require('google-spreadsheet');

module.exports = function(spreadsheetId, options) {
    var spreadsheet = new GoogleSpreadsheet(spreadsheetId);
    var callback = function(err) {
        throw err;
    };

    if (options) {
        if (isFunction(options)) {
            callback = options;
            options = {};
        } else if (isFunction(options.callback)) {
            callback = options.callback;
        } else if (isFunction(options.complete)) {
            callback = options.complete;
        }
    } else {
        options = {};
    }

    if (options.token) {
        spreadsheet.setAuthToken({
            value: options.token,
            type: options.tokentype || 'Bearer'
        });
        run(spreadsheet, callback, options);

    } else if (options.user && options.password) {
        spreadsheet.setAuth(options.user, options.password, function(err) {
            if (err)
                callback(err);

            run(spreadsheet, callback, options);
        });

    } else {
        run(spreadsheet, callback, options);

    }
}

function isFunction(func) {
    return !!(func && Object.prototype.toString.call(func) == '[object Function]');
}

function normalize(str) {
    return str
      .trim()
      .toLowerCase()
      .replace(/[- ]/ig, ' ')
      .split(' ')
      .map(function(val, index) {
          return !index ? val : val.charAt(0).toUpperCase() + val.slice(1);
      })
      .join('');
}

function run(spreadsheet, callback, options) {

    spreadsheet.getInfo(function(err, sheet_info) {
        if (err)
            callback(err);

        var json = {};
        var count = 0;

        sheet_info.worksheets.forEach(function(worksheet) {
            var title = normalize(worksheet.title);
            if (!options.worksheet ||
                !Array.isArray(options.worksheet) && normalize(options.worksheet) === title ||
                Array.isArray(options.worksheet) && options.worksheet.map(normalize).indexOf(title) !== -1
            ) {
                count++;
                worksheet.getCells(function(err, cells) {
                    if (err)
                        callback(err);

                    // setting up some options, such as defining if the data is horizontal or vertical

                    var rowProp = options.vertical ? 'col' : 'row';
                    var colProp = options.vertical ? 'row' : 'col';
                    var isHashed = options.hash && !options.listOnly;
                    var finalList = isHashed ? {} : [];

                    // organizing (and ordering) the cells into arrays

                    var rows = cells.reduce(function(rows, cell) {
                        var rowIndex = cell[rowProp] - 1;
                        if (typeof rows[rowIndex] === 'undefined')
                            rows[rowIndex] = [];
                        rows[rowIndex].push(cell);
                        return rows;
                    }, []);

                    rows.forEach(function(col) {
                        col.sort(function(cell1, cell2) {
                            return cell1[colProp] - cell2[colProp];
                        });
                    });

                    // find the first row with data to use it as property names

                    for (var firstRowIndex = 0; firstRowIndex < rows.length; firstRowIndex++) {
                        if (rows[firstRowIndex])
                            break;
                    }

                    // creating the property names map (to detect the name by index)

                    var properties = (rows[firstRowIndex] || []).reduce(function(properties, cell) {
                        if (typeof cell.value !== 'string' || cell.value === '')
                            return properties;

                        properties[cell[colProp]] = normalize(cell.value);

                        return properties;
                    }, {});

                    // removing first rows, before and including the one that is used as property names

                    rows.splice(0, firstRowIndex + 1);

                    // iterating through remaining row to fetch the values and build the final data object

                    rows.forEach(function(cells) {

                        var newObject = options.listOnly ? [] : {};
                        var hasValues = false;

                        cells.forEach(function(cell) {

                            var val;
                            var colNumber = cell[colProp];

                            if (!options.listOnly && !properties[colNumber])
                                return;

                            if (typeof cell.numericValue !== 'undefined') {
                                val = parseFloat(cell.numericValue);
                                hasValues = true;
                            } else if (cell.value === 'TRUE') {
                                val = true;
                                hasValues = true;
                            } else if (cell.value === 'FALSE') {
                                val = false;
                                hasValues = true;
                            } else if (cell.value !== '' && typeof cell.value !== 'undefined') {
                                val = cell.value;
                                hasValues = true;
                            } else {
                                return;
                            }

                            if (options.listOnly)
                                newObject[colNumber - 1] = val;
                            else
                                newObject[properties[colNumber]] = val;
                        });

                        if (hasValues) {
                            if (isHashed) {
                                finalList[newObject[options.hash]] = newObject;
                            } else {
                                finalList.push(newObject);
                            }
                        }
                    });

                    json[title] = finalList;
                    if (--count == 0)
                        complete(json, callback);

                });

            }

        });

        if (count == 0)
            complete(json, callback);
    });
}

function complete(json, callback) {
    var keys = Object.keys(json);
    if (keys.length === 1)
        json = json[keys[0]];

    callback(null, json);
}
