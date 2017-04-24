# dreamslate
A node.js command line tool to convert structured .json translation files to an Excel document and vice-versa.

## Installation
Clone the repo, then cd into it and run `npm install -g`.

## Usage
### Importing an excel file and generating translation files (dsimport)
Run `dsimport -i <importFileLocation> -o <destinationFolder>`
- *importFileLocation*: The excel file from which to generate the .json files
- *destinationFolder*: The output location to write the file(s) to

### Exporting translation files to an excel file
Run `dsexport -i <messagesFolder> -o <excelExportFile>`

- *messagesFolder*: The directory where the json file folders are located
- *excelExportFile*: The file name to export the excel file to

## Optional configuration file
Running these commands and constantly specifying parameters can be cumbersome. 
To run `dsimport` and `dsexport` without these parameters you can create an optional JSON configuration file named `translationConfig.json`.

```
{
	"excelExportFile": "mercury-translations.xlsx",
	"messagesFolder": "translate/",

	"importFileLocation": "mercury-translations.xlsx",
	"destinationFolder": "translations/"
}
```
The keys in this config files correspond to the values supplied in the above usages of the import and export commands.
 