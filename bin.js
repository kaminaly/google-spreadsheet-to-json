#!/usr/bin/env node

var fs = require('fs');
var program = require('commander');
var packageData = require('./package.json');
var gsjson = require('./index');

program
    .version(packageData.version)
    .usage('<spreadsheet-id> <file> [options]')
    .option('-u, --user [user]', 'User to login')
    .option('-p, --password [password]', 'Password to login')
    .option('-t, --token [token]', 'Auth token acquired externally')
    .option('-y, --tokentype [tokentype]', 'Type of the informed token (defaults to Bearer)')
    .option('-w, --worksheet <sheetname[,sheetname]>', 'Worksheet name from Google tab UI (defaults to all worksheet)')
    .option('-c, --hash [column]', 'Column to hash the final JSON')
    .option('-i, --vertical', 'Use the first column as header')
    .option('-l, --list-only', 'Ignore headers and just list the values in arrays')
    .option('-b, --beautify', 'Beautify final JSON')
    .parse(process.argv);

if (program.args.length < 2) {
    program.help();
}

if(program.worksheet && program.worksheet.indexOf(',') !== -1) {
    program.worksheet = program.worksheet.split(',').map(function(v){return v.trim()});
}

var spreadsheetId = program.args[0];
var filename = program.args[1];

program.callback = function(err, json){
    if (err)
        throw err;

    fs.writeFile(filename, json, 'utf-8', function(err) {
        if (err)
            throw err;
        process.exit(0);
    });
}

gsjson(spreadsheetId, program);
