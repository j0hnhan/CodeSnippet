/*
	Self written code to extract data from .xlsx and replace 
	it into text template
	Written by JavaScript, running in Node.js
	Work as a lightweight simple script
 */
var fs = require('fs');
var Q = require('q');
var textract = require('textract');
var pdfkit = require('pdfkit');

/**
 * Read .docx file content and generate a promise
 * @param  {String}		FileName of the .docx file
 * @return {Promise} 	Returns a promise for further use
 */
function getDocxPattern(fileName) {
	return function() {
		var myDefer = Q.defer();
		textract.fromFileWithPath(fileName, function(err, text) {
			if(!err && text) {
				myDefer.resolve(text.trim());
			} else if(err) {
				console.log("Error: " + err);
			}
		});

		return myDefer.promise;
	}
}

/**
 * Read .xlsx file content and generate a promise
 * @param  {String} 	Filename of the .xlsx file
 * @param  {String}		Content retrieved from .docx file
 * @return {Pormise}	Returns a promise for further process
 */
function getXlsx(fileName, pattern) {
	return function() {
		var myDefer = Q.defer();
		textract.fromFileWithPath(fileName, function(err, text) {
			if(!err && text) {
				var records = text.trim().split(" ");
				myDefer.resolve({
					pattern: pattern,
					records: records
				});
			} else if(err) {
				console.log("Error: " + err);
			}
		});

		return myDefer.promise;
	}
}


/**
 * Generate pdf file by replace pattern's name and date by record
 * @param  {String}		one line record of .xlxs file
 * @param  {String}		String need to be replaced
 * @param  {Number}		index indicate pdf index.
 * @return {undefined}
 */
function generatePdf(record, pattern, idx) {
	var replacedText = pattern.replace('NAME', record.split(",")[0])
				.replace('DATE', record.split(",")[1]);
	var doc = new pdfkit();
	doc.pipe(fs.createWriteStream('output' + idx + '.pdf'));
	doc.font('Times-Roman')
		.text(replacedText);
	doc.end();
}

getDocxPattern("test.docx")().then(function(value) {
	return getXlsx("test.xlsx", value)();
}).then(function(value) {
	for(var i = 0; i < value.records.length; i++) {
		generatePdf(value.records[i], value.pattern, i);
	}
}).done();

