#!/usr/bin/env node

"use strict";
const argv = require("minimist")(process.argv.slice(2));
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const expressions = require("angular-expressions");
const csvtojson=require('csvtojson');

function transformError(error) {
	const e = {
		message: error.message,
		name: error.name,
		stack: error.stack,
		properties: error.properties,
	};
	if (e.properties && e.properties.rootError) {
		e.properties.rootError = transformError(error.properties.rootError);
	}
	if (e.properties && e.properties.errors) {
		e.properties.errors = e.properties.errors.map(transformError);
	}
	return e;
}

function printErrorAndRethrow(error) {
	const e = transformError(error);
	// The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
	console.error(JSON.stringify({ error: e }, null, 2));
	throw error;
}

function showHelp() {
	console.log("Usage: docxtemplater input.docx data.json output.docx");
	process.exit(1);
}

function parser(tag) {
	const expr = expressions.compile(tag.replace(/’/g, "'"));
	return {
		get(scope) {
			return expr(scope);
		},
	};
}

function render(doc,data,outputFile) {
	doc.setData(data);

	try {
		doc.render();
	} catch (error) {
		printErrorAndRethrow();
	}

	const generated = doc
		.getZip()
		.generate({ type: "nodebuffer", compression: "DEFLATE" });

	fs.writeFileSync(outputFile, generated);
}

const args = argv._;
if (argv.help || args.length !== 3) {
	showHelp();
}
let options = {};
if (argv.options) {
	try {
		options = JSON.parse(argv.options);
	} catch (e) {
		console.error("Arguments passed in --options is not valid JSON");
		throw e;
	}
}

options.parser = parser;

async function main() {

	const [inputFile, dataFile, outputFile] = args;
	const input = fs.readFileSync(inputFile, "binary");
	let data;
	if (dataFile.endsWith(".json")){
		data = JSON.parse(fs.readFileSync(dataFile, "utf-8"));
	} else if (dataFile.endsWith(".csv")){
		data = await csvtojson({flatKeys: true}).fromFile(dataFile);
	} else {
		throw new Error(`Unsupported file format ${dataFile}`);
	}


	let doc;

	try {
		doc = new Docxtemplater(new PizZip(input), options);
	} catch (e) {
		printErrorAndRethrow(e);
	}

	if (Array.isArray(data)){
		let i = 0;
		for(const elem of data){
			i++;
			render(doc,elem,elem['_filename'] ?? outputFile.replace('.docx',`${i}.docx`));
		}
	} else {
		render(doc,data,outputFile);
	}

}

main()
	.then(() => process.exit(0))
	.catch(e => {console.error(e); process.exit(1);})
;


