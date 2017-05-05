#!/usr/bin/env node
var program = require('commander');
var Excel = require("exceljs");
var fs = require('fs');
var gutil = require('gulp-util');
var rename = require('gulp-rename');
var replace = require('gulp-replace');
var	gulp = require('gulp');
var jsonfile = require('jsonfile');
var _ = require('underscore');
var configuration = require('./config.js');
var styles = configuration.styles;
var config = configuration.config;

// cli
program
	.option('-o, --exportFile <excelExportFile>', 'The file name to export the excel file to')
	.option('-i, --messages <messagesFolder>', 'The directory where the json files are located')
	.option('-n, --messages <numberOfLanguages>', 'The number of languages you want to translate')
	.parse(process.argv);

// if a file has been specified use it here
var optionalConfigFile;
try {
	optionalConfigFile = jsonfile.readFileSync(configuration.optionalConfigFileName)
}catch(e) {}

if (!optionalConfigFile && (!program.exportFile || !program.messages)) {
	gutil.log(gutil.colors.bgRed('missing required arguments ... run --help'));
	return;
}

var paths = {
	excelExport: program.excelExportFile || optionalConfigFile.excelExportFile, //'mercury-translations.xlsx',
	messagesDir: './' + (program.messagesFolder || optionalConfigFile.messagesFolder) ,//'./Client/messages/'
};

var numOfLangs = (program.numberOfLanguages || optionalConfigFile.numberOfLanguages) || 2;

function writeXLSX(callback){

	// instead of hardcoding the languages we can infer them from the folder names in the messages folder
	var languages = fs.readdirSync(paths.messagesDir);

	// filter out any hidden folders like .DS_STORE
	languages = _.filter(languages, function(key) {return key.charAt(0) !== '.'});

	// add a new excel workbook that will eventually be exported
	var workbook = new Excel.Workbook();

	// adding worksheets to the excel document
	var worksheets = [];

	var columns = config.baseColumns;
	var translations = {};

	for (var i = 0; i < languages.length; i++) {
		var language = languages[i];
		translations[language] = {};

		// reading all the names in language directory
		var thisTranslationDirectory = paths.messagesDir + language + '/';

		var translationFiles = fs.readdirSync(thisTranslationDirectory);
		// filter out any hidden files names like .DS_STORE
		translationFiles = _.filter(translationFiles, function(key) {return key.charAt(0) !== '.'});

		for(var fileIndex = 0 ; fileIndex < translationFiles.length; fileIndex++){
			var translationFile = translationFiles[fileIndex]; // HOME_PAGE.json
			var translationFileKey = translationFile.split(".")[0]; // HOME_PAGE

			var filepath = thisTranslationDirectory + translationFile;
			translations[language] = jsonfile.readFileSync(filepath, null);

			var pageNames = Object.keys(translations[language]);
			// add the name of this key to the worksheets
			worksheets = _.union(worksheets, pageNames);
		}

		columns.push({
			header: language,
			key: language,
			width: 50,
			style: styles.cellTranslation
		});
	}

	for (var i in worksheets) {
		workbook.addWorksheet(worksheets[i]);
	}
	
	var duplicatedKeys = [];

	function addRow(worksheet, page, section, key){
		var row = {page: page, section: section, key: key};
		var item = translations[page][section][key];

		// alter the row data if this is a less nested item
		if (typeof item === "string"){
			item = translations[page][section];
			row.section = "";
			row.key = section;
		}

		for (var language in item) {
			row[language] = item[language];
		}
		worksheet.addRow(row);
	}

	workbook.eachSheet(function(worksheet, sheetId) {
		worksheet.columns = columns;
		worksheet.getRow(1).font = styles.header.font;
		worksheet.getRow(1).fill = styles.header.fill;
		worksheet.getRow(1).border = styles.header.border;
	});

	// convert the structure from
	// language
	//     |_ page
	//          |_ section
	//                 |_key
	//                     |_ value
	//
	// to...
	//
	//  page
	//   |_ section
	//         |_key
	//             |_language
	//                  |_ value

	var tmp = {};
	for (var language in translations) {
		for (var page in translations[language]) {
			tmp[page] = tmp[page] || {};
			var worksheet = workbook.getWorksheet(page);
			for (var section in translations[language][page]) {
				tmp[page][section] = tmp[page][section] || {};

				// this is in the event that the .json is two levels deep (i.e. there is no real section)
				if (typeof translations[language][page][section] === "string"){
					tmp[page][section][language] = translations[language][page][section];
					continue;
				}

				for (var key in translations[language][page][section]){
					tmp[page][section][key] = tmp[page][section][key] || {};
					tmp[page][section][key][language] = translations[language][page][section][key];

				}
			}
		}
	}

	translations = tmp;

	for (var page in translations) {
		var worksheet = workbook.getWorksheet(page);
		for (var section in translations[page]) {
			for (var key in translations[page][section]){
				if (duplicatedKeys.indexOf(section) === -1) {
					addRow(worksheet, page, section, key);
					if (typeof translations[page][section][key] === "string") {
						duplicatedKeys.push(section);
					}	
				}
			}
		}
	}

	// color rows that have missing information
	workbook.eachSheet(function(worksheet, sheetId) {
		worksheet.eachRow(function(row, rowNumber) {
			for (var i = columns.length - numOfLangs; i <= columns.length; i++){
				if (!row.values[i]){
					worksheet.getRow(rowNumber).fill = styles.emptyRow.fill;
					break;
				}
			}
		});
	});

	workbook.xlsx.writeFile(paths.excelExport).then(function(){
		callback(paths.excelExport);
	});
}

writeXLSX(function(translation){
	gutil.log(gutil.colors.green('Success') +' - Translations exported to ' + gutil.colors.magenta('~'+paths.excelExport));
});

module.exports = {
	writeXLSX : writeXLSX
};
