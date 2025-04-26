const fs = require('fs');
const JSZip = require('jszip');

/**
 * Processes tables in Word documents to add serial numbers in REF columns
 * @param {string} outputPath - Path to the merged document
 * @returns {Promise<void>}
 */
async function processTablesInDocument(outputPath) {
  try {
    console.log(`Processing tables in document: ${outputPath}`);
    
    // Read the document file
    const documentContent = fs.readFileSync(outputPath);
    const zip = await JSZip.loadAsync(documentContent);
    
    // Get the main document content
    const documentXml = await zip.file('word/document.xml').async('string');
    
    // Process the document content
    let serialNumber = 1;
    const processedXml = processTableReferences(documentXml, serialNumber);
    
    // Save the processed content back to the document
    zip.file('word/document.xml', processedXml);
    
    // Generate the new document
    const content = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE'
    });
    
    // Write the updated document
    fs.writeFileSync(outputPath, content);
    
    console.log('Table REF columns processed successfully');
  } catch (error) {
    console.error('Error processing tables:', error);
    throw error;
  }
}

/**
 * Process the document XML to find and number tables
 * @param {string} docXml - The document XML content
 * @param {number} startNumber - The starting serial number
 * @returns {string} - The processed XML with numbered tables
 */
function processTableReferences(docXml, startNumber) {
  let serialNumber = startNumber;
  console.log(`Starting with serial number: ${serialNumber}`);
  
  // Extract all tables from the document
  const tableRegex = /<w:tbl>[\s\S]*?<\/w:tbl>/g;
  const tables = docXml.match(tableRegex);
  
  if (!tables || tables.length === 0) {
    console.log("No tables found in the document.");
    return docXml;
  }
  
  console.log(`Found ${tables.length} tables in the document.`);
  
  // Process each table
  let processedXml = docXml;
  for (let i = 0; i < tables.length; i++) {
    const originalTable = tables[i];
    const processedTable = processTable(originalTable, i + 1, serialNumber);
    
    // Update the serialNumber based on how many rows were processed
    const processedRows = (processedTable.match(/<REF_NUMBER_ADDED>/g) || []).length;
    serialNumber += processedRows;
    
    // Replace the original table with the processed one
    processedXml = processedXml.replace(originalTable, processedTable.replace(/<REF_NUMBER_ADDED>/g, ''));
  }
  
  console.log(`Completed numbering, last serial number: ${serialNumber - 1}`);
  return processedXml;
}

/**
 * Process an individual table to add serial numbers to the REF column
 * @param {string} tableXml - The table XML
 * @param {number} tableIndex - The index of the table
 * @param {number} startSerialNumber - The starting serial number
 * @returns {string} - The processed table XML
 */
function processTable(tableXml, tableIndex, startSerialNumber) {
  console.log(`Processing table ${tableIndex}`);
  
  // Extract rows from the table
  const rowRegex = /<w:tr[\s\S]*?<\/w:tr>/g;
  const rows = tableXml.match(rowRegex);
  
  if (!rows || rows.length === 0) {
    console.log(`Table ${tableIndex} has no rows.`);
    return tableXml;
  }
  
  console.log(`Table ${tableIndex} has ${rows.length} rows.`);
  
  let serialNumber = startSerialNumber;
  let processedTableXml = tableXml;
  
  // Process each row
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    
    // Skip header rows: first row (index 0) and rows 5, 9, 13, etc.
    const rowNumber = rowIndex + 1;
    if (rowNumber === 1 || (rowNumber % 4 === 1 && rowNumber > 1)) {
      console.log(`  Skipping header row ${rowNumber}`);
      continue;
    }
    
    // Check if this is a row we should process
    const processedRow = processRowFirstCell(row, serialNumber);
    
    if (processedRow !== row) {
      // Replace the row in the table
      processedTableXml = processedTableXml.replace(row, processedRow);
      serialNumber++;
    }
  }
  
  return processedTableXml;
}

/**
 * Process the first cell of a row to add a serial number
 * @param {string} rowXml - The row XML
 * @param {number} serialNumber - The serial number to add
 * @returns {string} - The processed row XML
 */
function processRowFirstCell(rowXml, serialNumber) {
  // Find the first cell in the row
  const cellRegex = /<w:tc>[\s\S]*?<\/w:tc>/;
  const cellMatch = rowXml.match(cellRegex);
  
  if (!cellMatch) {
    return rowXml;
  }
  
  const cell = cellMatch[0];
  
  // Find the paragraph in the cell
  const paragraphRegex = /<w:p[\s\S]*?<\/w:p>/;
  const paragraphMatch = cell.match(paragraphRegex);
  
  if (!paragraphMatch) {
    // No paragraph found, insert new one with serial number
    const newCell = cell.replace(/<w:tc>([\s\S]*?)/, 
      `<w:tc>$1<w:p><w:r><w:t>${serialNumber}</w:t></w:r></w:p><REF_NUMBER_ADDED>`);
    return rowXml.replace(cell, newCell);
  }
  
  // Replace text in the paragraph with the serial number
  const paragraph = paragraphMatch[0];
  
  // Create a new paragraph with the serial number
  let newParagraph;
  
  // Keep paragraph properties if they exist
  const pPrRegex = /<w:pPr>[\s\S]*?<\/w:pPr>/;
  const pPrMatch = paragraph.match(pPrRegex);
  
  if (pPrMatch) {
    newParagraph = paragraph.replace(
      pPrMatch[0] + '[\s\S]*?<\/w:p>',
      `${pPrMatch[0]}<w:r><w:t>${serialNumber}</w:t></w:r></w:p><REF_NUMBER_ADDED>`
    );
  } else {
    newParagraph = `<w:p><w:r><w:t>${serialNumber}</w:t></w:r></w:p><REF_NUMBER_ADDED>`;
  }
  
  // Replace the paragraph in the cell
  const newCell = cell.replace(paragraph, newParagraph);
  
  // Replace the cell in the row
  return rowXml.replace(cell, newCell);
}

module.exports = {
  processTablesInDocument
}; 