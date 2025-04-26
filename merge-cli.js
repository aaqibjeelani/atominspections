#!/usr/bin/env node

const path = require('path');
const fs = require('fs');
const { mergeDocuments } = require('./wordMerger');

// Check if we have enough arguments
if (process.argv.length < 4) {
  console.error(`
Usage: node merge-cli.js [output-file] [input-file1] [input-file2] ...

Example: node merge-cli.js merged.docx document1.docx document2.docx

The template file is automatically used from the sample/template.docx location.
  `);
  process.exit(1);
}

// Get arguments
const outputFilePath = process.argv[2];
const inputFilePaths = process.argv.slice(3);

// Validate input files
inputFilePaths.forEach(filePath => {
  if (!fs.existsSync(filePath)) {
    console.error(`Error: Input file does not exist: ${filePath}`);
    process.exit(1);
  }
  
  if (!filePath.endsWith('.docx')) {
    console.error(`Error: Input file is not a .docx file: ${filePath}`);
    process.exit(1);
  }
});

// Template path
const templatePath = path.join(__dirname, 'sample', 'template.docx');

// Validate template
if (!fs.existsSync(templatePath)) {
  console.error(`Error: Template file does not exist at ${templatePath}`);
  process.exit(1);
}

// Merge the documents
console.log('Merging documents...');
console.log(`Template: ${templatePath}`);
console.log(`Input files: ${inputFilePaths.join(', ')}`);
console.log(`Output: ${outputFilePath}`);

mergeDocuments(templatePath, inputFilePaths, outputFilePath)
  .then(() => {
    console.log('Documents merged successfully!');
    console.log(`Output file saved to: ${outputFilePath}`);
  })
  .catch(error => {
    console.error('Error merging documents:', error);
    process.exit(1);
  }); 