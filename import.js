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
	.option('-i, --import <importFile>', 'Imports an excel file and generates .json translation files')
	.option('--single-files', 'When importing, should all translations for a given language go in one file or be split up based on their page')
	.option('-o, --output <destination>', 'The output location to write the file(s) to')
	.parse(process.argv);



var optionalConfigFile;
try {
	optionalConfigFile = jsonfile.readFileSync(configuration.optionalConfigFileName)
}catch(e) {}

if (!optionalConfigFile && (!program.import || !program.output)) {
	console.log('missing required arguments ... run --help');
	return;
}

var paths = {
	excelImport: program.import || optionalConfigFile.excelImport,//'mercury-translations.xlsx',
	translationsDir : program.output || optionalConfigFile.translationsDir,//'translations/',
	translationTemplate: 'translation-template.json'
};

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
			for (var section in translation[lang]) {
				var filename = lang +'/' + section + '.json';
				gulp.src(paths.translationTemplate)
					.pipe(replace(/\/\*js-inject:translations\*\//g, JSON.stringify(translation[lang][section], null, 4)))
					.pipe(rename(filename))
					.pipe(gulp.dest(paths.translationsDir));
				gutil.log(gutil.colors.green('Success') + ' - Translations imported ' + gutil.colors.magenta('~'+paths.translationsDir+filename));
			}
		}
	}else{
		gutil.log(gutil.colors.red('Translation tables not created'));
	}
});

module.exports = {
	readXLSX : readXLSX
};
