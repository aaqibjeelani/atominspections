const fs = require('fs');
const path = require('path');
const DocxMerger = require('docx-merger');
const { processTablesInDocument } = require('./tableHandler');

/**
 * Merges multiple Word documents using a template document
 * @param {string} templatePath - Path to the template .docx file
 * @param {Array<string>} documentPaths - Array of paths to document files to merge
 * @param {string} outputPath - Path where the merged document will be saved
 * @returns {Promise<Buffer>} - Buffer containing the merged document
 */
async function mergeDocuments(templatePath, documentPaths, outputPath) {
  return new Promise((resolve, reject) => {
    try {
      console.log('-------- Starting document merge --------');
      console.log(`Template file: ${templatePath}`);
      console.log(`Number of documents to merge: ${documentPaths.length}`);
      console.log(`Output path: ${outputPath}`);
      
      // Check if template exists
      if (!fs.existsSync(templatePath)) {
        throw new Error(`Template file does not exist: ${templatePath}`);
      }
      
      // Check if all documents exist
      for (const docPath of documentPaths) {
        if (!fs.existsSync(docPath)) {
          throw new Error(`Document file does not exist: ${docPath}`);
        }
      }
      
      // Read all document files to merge
      console.log('Reading document files...');
      const docxBuffers = documentPaths.map((docPath, index) => {
        console.log(`  Reading document ${index + 1}: ${path.basename(docPath)}`);
        return fs.readFileSync(docPath);
      });
      
      // Read the template document
      console.log('Reading template document...');
      const templateBuffer = fs.readFileSync(templatePath);
      
      // Create a separate array for documents only
      const mergeBuffers = [...docxBuffers];
      
      // Create the DocxMerger instance
      console.log('Merging documents...');
      const docxMerger = new DocxMerger({
        pageBreak: true,    // Add page breaks between documents
        continuous: true,   // Make sections continuous (avoid blank pages)
      }, mergeBuffers);
      
      // Save the merged document to a temporary file
      const tempOutputPath = outputPath + '.temp';
      console.log(`Saving merged content to temporary file: ${tempOutputPath}`);
      
      docxMerger.save('nodebuffer', async (mergedBuffer) => {
        try {
          // Write the merged content to a temp file
          fs.writeFileSync(tempOutputPath, mergedBuffer);
          console.log('Merge complete. Starting template application...');
          
          // Apply the template formatting
          const { applyTemplate } = require('./templateApplier');
          await applyTemplate(templatePath, tempOutputPath, outputPath);
          
          console.log('Template applied. Processing table numbering...');
          
          // Process tables to add REF column numbering
          await processTablesInDocument(outputPath);
          
          // Clean up the temporary file
          console.log('Cleaning up temporary files...');
          if (fs.existsSync(tempOutputPath)) {
            fs.unlinkSync(tempOutputPath);
          }
          
          console.log('-------- Document merge completed --------');
          // Return the final file path
          resolve(outputPath);
        } catch (error) {
          console.error('Error in merge process:', error);
          reject(error);
        }
      });
    } catch (error) {
      console.error('Error merging documents:', error);
      reject(error);
    }
  });
}

/**
 * Cleans up temporary files created during the merging process
 * @param {Array<string>} filePaths - Array of file paths to delete
 */
function cleanupFiles(filePaths) {
  console.log('Cleaning up files...');
  filePaths.forEach(filePath => {
    if (fs.existsSync(filePath)) {
      try {
        fs.unlinkSync(filePath);
        console.log(`  Deleted file: ${filePath}`);
      } catch (error) {
        console.error(`  Error deleting file ${filePath}:`, error);
      }
    } else {
      console.log(`  File doesn't exist, skipping: ${filePath}`);
    }
  });
}

module.exports = {
  mergeDocuments,
  cleanupFiles
}; 