var program = require('commander');
var Excel = require("exceljs");
var fs = require('fs');
var gutil = require('gulp-util');
var rename = require('gulp-rename');
var replace = require('gulp-replace');
var	gulp = require('gulp');
var jsonfile = require('jsonfile');
var configuration = require('./config.js');

// cli
program
	.option('-i, --import <importFileLocation>', 'Imports an excel file and generates .json translation files')
	.option('--single-files [singleFilesName]', 'When importing, should all translations for a given language go in one file or be split up based on their page')
	.option('-o, --output <destinationFolder>', 'The output location to write the file(s) to')
	.parse(process.argv);



var optionalConfigFile;
try {
	optionalConfigFile = jsonfile.readFileSync(configuration.optionalConfigFileName)
}catch(e) {}

if (!optionalConfigFile && (!program.import || !program.output)) {
	gutil.log(gutil.colors.bgRed('missing required arguments ... run --help'));
	return;
}

var paths = {
	excelImport: program.import || optionalConfigFile.importFileLocation,//'mercury-translations.xlsx',
	translationsDir : program.output || optionalConfigFile.destinationFolder,//'translations/',
	translationTemplate: __dirname+'/translation-template.json'
};

var importAsSingleFile = program.singleFiles || optionalConfigFile.importFromExcelAsSingleFiles || false;
var singleFilesName = program.singleFilesName || optionalConfigFile.singleFilesName || 'imported-translation.json';

function readXLSX(callback){
	var workbook = new Excel.Workbook();
	var translations = {};

	fs.exists(paths.excelImport, function(exists) {
		if (!exists) {
			gutil.log(gutil.colors.red('Error - No File found : ' + paths.excelImport));
			return;
		}

		// read the excel file
		workbook.xlsx.readFile(paths.excelImport)
			.then(function() {

				// loop through each sheet, adding the languages
				workbook.eachSheet(function(worksheet, sheetId) {
					var languages = getLanguagesForWorksheet(worksheet, 4);
					var page = worksheet.name;

					for(var language in languages){
						translations[language] = translations[language] || {};
						translations[language][page] = {};
					}

					for(var row = 2; row <= worksheet.lastRow.number; row++){
						var section = worksheet.getRow(row).getCell(2).value;
						var key = worksheet.getRow(row).getCell(3).value;

						// loop through the languages to get the translations for this row
						for(var language in languages){
							var translation = worksheet.getRow(row).getCell(languages[language]).value;

							if(translation === null || translation === ""){
								translation = "";
							}

							if(section){
								if(!(section in translations[language][page])){
									translations[language][page][section] = {};
								}
								translations[language][page][section][key] = translation;
							}else{
								translations[language][page][key] = translation;
							}
						}
					}
				});
				callback(translations);
			});
	});
}

/**
 * Populates an object with the languages and their respective column numbers in a given worksheet
 * @param worksheet - the excel worksheet to get the languages from
 * @param startingLanguageColumn - the first column in the first row to start recording languages and their column numbers from
 * @returns {object} - an object where the keys are the languages and the values are the column numbers where the languages occur
 */
function getLanguagesForWorksheet(worksheet, startingLanguageColumn) {
	var output = {};
	worksheet.getRow(1).eachCell(function (cell, colNumber) {
		if (colNumber >= startingLanguageColumn) {
			output[cell.value] = colNumber;
		}
	});
	return output;
}

readXLSX(function(translation){
	if(translation){
		for(var lang in translation){
			if (importAsSingleFile){
				var filename = lang + '/' + singleFilesName;
				writeJSONToFile(filename, translation[lang]);
				continue;
			}

			for (var page in translation[lang]) {
				var thisSection = {};
				thisSection[page] = translation[lang][page];
				var filename = lang + '/' + page + '.json';
				writeJSONToFile(filename, thisSection);
			}
		}
	}else{
		gutil.log(gutil.colors.red('Translation tables not created'));
	}
});

function writeJSONToFile(filename, jsonData) {
	gulp.src(paths.translationTemplate)
		.pipe(replace(/\/\*js-inject:translations\*\//g, JSON.stringify(jsonData, null, 4)))
		.pipe(rename(filename))
		.pipe(gulp.dest(paths.translationsDir));
	gutil.log(gutil.colors.green('Success') + ' - Translations imported ' + gutil.colors.magenta(paths.translationsDir+filename));
}

module.exports = {
	readXLSX : readXLSX
};
