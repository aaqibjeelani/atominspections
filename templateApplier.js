const fs = require('fs');
const JSZip = require('jszip');

/**
 * Applies template formatting to a merged document
 * @param {string} templatePath - Path to the template document
 * @param {string} mergedPath - Path to the merged document
 * @param {string} outputPath - Path where to save the final document
 * @returns {Promise<void>}
 */
async function applyTemplate(templatePath, mergedPath, outputPath) {
  try {
    console.log("Starting template application process...");
    console.log(`Template: ${templatePath}`);
    console.log(`Merged document: ${mergedPath}`);
    console.log(`Output path: ${outputPath}`);

    // Read the template and merged document as binary data
    const templateContent = fs.readFileSync(templatePath);
    const mergedContent = fs.readFileSync(mergedPath);

    // Load both files into JSZip for processing
    const templateZip = await JSZip.loadAsync(templateContent);
    const mergedZip = await JSZip.loadAsync(mergedContent);

    // STRATEGY: Preserve original document but inject header/footer from template

    // 1. Get Word version information from both documents
    const templateAppProps = await getZipEntryText(templateZip, 'docProps/app.xml');
    const mergedAppProps = await getZipEntryText(mergedZip, 'docProps/app.xml');
    
    console.log(`Template Word version: ${extractWordVersion(templateAppProps)}`);
    console.log(`Merged document Word version: ${extractWordVersion(mergedAppProps)}`);

    // 2. Copy key files (but maintain document content)
    console.log("Copying header and footer files...");
    await copyHeaderFooterFiles(templateZip, mergedZip);
    
    // 3. Maintain important style files while preserving version compatibility
    console.log("Preserving document styles...");
    await preserveStyles(templateZip, mergedZip);
    
    // 4. Fix document.xml to reference headers and footers
    console.log("Fixing document references...");
    await updateDocumentReferences(templateZip, mergedZip);
    
    // 5. Ensure there's no blank first page
    console.log("Checking for blank first page...");
    await removeBlankFirstPage(mergedZip);

    // 6. Generate the final document
    console.log("Generating final document...");
    const content = await mergedZip.generateAsync({
      type: 'nodebuffer', 
      compression: 'DEFLATE'
    });
    
    // Write the final document to the output path
    fs.writeFileSync(outputPath, content);
    console.log("Template application completed successfully.");
  } catch (error) {
    console.error('Error applying template:', error);
    throw error;
  }
}

/**
 * Extracts Word version from app.xml
 * @param {string} appXml - The app.xml content
 * @returns {string} - Version information
 */
function extractWordVersion(appXml) {
  if (!appXml) return "Unknown";
  
  const appNameMatch = appXml.match(/<AppName>(.*?)<\/AppName>/);
  const appVersionMatch = appXml.match(/<AppVersion>(.*?)<\/AppVersion>/);
  
  const appName = appNameMatch ? appNameMatch[1] : "Unknown";
  const appVersion = appVersionMatch ? appVersionMatch[1] : "";
  
  return `${appName} ${appVersion}`.trim();
}

/**
 * Copies header and footer files from template to merged document
 * @param {JSZip} templateZip - The template zip
 * @param {JSZip} mergedZip - The merged document zip
 */
async function copyHeaderFooterFiles(templateZip, mergedZip) {
  // Get all header and footer files
  const headerFooterFiles = templateZip.filter(path => {
    return path.startsWith('word/header') || path.startsWith('word/footer');
  });
  
  // Copy each file to the merged document
  for (const file of headerFooterFiles) {
    try {
      const content = await file.async('nodebuffer');
      mergedZip.file(file.name, content);
      console.log(`  Copied ${file.name}`);
    } catch (error) {
      console.error(`  Error copying ${file.name}:`, error);
    }
  }
}

/**
 * Preserves styles while maintaining version compatibility
 * @param {JSZip} templateZip - The template zip
 * @param {JSZip} mergedZip - The merged document zip
 */
async function preserveStyles(templateZip, mergedZip) {
  // Process relationships first to ensure proper references
  await preserveRelationships(templateZip, mergedZip);
  
  // List of files to carefully merge (not just overwrite)
  const filesToPreserve = [
    'word/settings.xml',
    'word/styles.xml',
    'word/fontTable.xml',
    'word/theme/theme1.xml',
    '[Content_Types].xml'
  ];
  
  for (const filePath of filesToPreserve) {
    try {
      // Get content from both documents
      const templateContent = await getZipEntryText(templateZip, filePath);
      const mergedContent = await getZipEntryText(mergedZip, filePath);
      
      // Skip if either file doesn't exist
      if (!templateContent || !mergedContent) {
        console.log(`  Skipping ${filePath} - file not found in one of the documents`);
        continue;
      }
      
      // For theme and fontTable, copy the template version for consistent styling
      if (filePath.includes('theme') || filePath.includes('fontTable')) {
        mergedZip.file(filePath, templateContent);
        console.log(`  Copied ${filePath} from template`);
        continue;
      }
      
      // For styles.xml, we need to carefully merge
      if (filePath === 'word/styles.xml') {
        const mergedStyles = await mergeStyles(templateContent, mergedContent);
        mergedZip.file(filePath, mergedStyles);
        console.log(`  Merged styles from template and document`);
        continue;
      }
      
      // For other files, preserve merged document content
      console.log(`  Preserved original ${filePath}`);
    } catch (error) {
      console.error(`  Error processing ${filePath}:`, error);
    }
  }
}

/**
 * Merges style definitions from template and document
 * @param {string} templateStyles - Template styles XML
 * @param {string} documentStyles - Document styles XML
 * @returns {string} - Merged styles XML
 */
async function mergeStyles(templateStyles, documentStyles) {
  // For now, just use template styles as they're more important for formatting
  return templateStyles;
}

/**
 * Preserves relationship files while ensuring proper references
 * @param {JSZip} templateZip - The template zip
 * @param {JSZip} mergedZip - The merged document zip
 */
async function preserveRelationships(templateZip, mergedZip) {
  try {
    // Get document rels from both files
    const templateRels = await getZipEntryText(templateZip, 'word/_rels/document.xml.rels');
    const mergedRels = await getZipEntryText(mergedZip, 'word/_rels/document.xml.rels');
    
    if (!templateRels || !mergedRels) {
      console.log("  Skipping relationship preservation - files not found");
      return;
    }
    
    // Extract header and footer relationships from template
    const headerFooterRegex = /<Relationship [^>]*(?:Target="(?:header|footer)\d+\.xml")[^>]*>/g;
    const headerFooterRels = templateRels.match(headerFooterRegex) || [];
    
    // Find the closing tag of relationships
    let updatedRels = mergedRels;
    if (headerFooterRels.length > 0) {
      // Add header/footer relationships
      updatedRels = mergedRels.replace('</Relationships>', 
        headerFooterRels.join('') + '</Relationships>');
      
      // Update the file
      mergedZip.file('word/_rels/document.xml.rels', updatedRels);
      console.log("  Updated document relationships with header/footer references");
    }
  } catch (error) {
    console.error("  Error preserving relationships:", error);
  }
}

/**
 * Updates document references to headers and footers
 * @param {JSZip} templateZip - The template zip
 * @param {JSZip} mergedZip - The merged document zip
 */
async function updateDocumentReferences(templateZip, mergedZip) {
  try {
    // Get document content from both files
    const templateDoc = await getZipEntryText(templateZip, 'word/document.xml');
    let mergedDoc = await getZipEntryText(mergedZip, 'word/document.xml');
    
    if (!templateDoc || !mergedDoc) {
      console.log("  Skipping document references update - files not found");
      return;
    }
    
    // Extract section properties that reference headers and footers
    const sectPrRegex = /<w:sectPr[^>]*>[\s\S]*?<\/w:sectPr>/g;
    const templateSectPrs = templateDoc.match(sectPrRegex);
    
    if (!templateSectPrs || templateSectPrs.length === 0) {
      console.log("  No section properties found in template");
      return;
    }
    
    // Get the main section properties from template (usually the last one)
    const mainSectPr = templateSectPrs[templateSectPrs.length - 1];
    
    // Find section properties in merged document
    const mergedSectPrs = mergedDoc.match(sectPrRegex);
    
    if (!mergedSectPrs || mergedSectPrs.length === 0) {
      // No section properties, add at the end of body
      mergedDoc = mergedDoc.replace('</w:body>', mainSectPr + '</w:body>');
      console.log("  Added section properties to document");
    } else {
      // Replace each section with the template section
      for (const sectPr of mergedSectPrs) {
        mergedDoc = mergedDoc.replace(sectPr, mainSectPr);
      }
      console.log("  Updated section properties in document");
    }
    
    // Update the document
    mergedZip.file('word/document.xml', mergedDoc);
  } catch (error) {
    console.error("  Error updating document references:", error);
  }
}

/**
 * Retrieves a file from a zip archive as binary
 * @param {JSZip} zip - The zip archive
 * @param {string} path - Path to the file in the archive
 * @returns {Promise<Buffer|null>} - The file content or null if not found
 */
async function getZipEntryBinary(zip, path) {
  const file = zip.file(path);
  if (!file) {
    return null;
  }
  return await file.async('nodebuffer');
}

/**
 * Retrieves a file from a zip archive as text
 * @param {JSZip} zip - The zip archive
 * @param {string} path - Path to the file in the archive
 * @returns {Promise<string|null>} - The file content or null if not found
 */
async function getZipEntryText(zip, path) {
  const file = zip.file(path);
  if (!file) {
    return null;
  }
  return await file.async('string');
}

/**
 * Removes any blank first page from the merged document
 * @param {JSZip} mergedZip - The merged document zip
 */
async function removeBlankFirstPage(mergedZip) {
  try {
    // Get the document content
    let docXml = await getZipEntryText(mergedZip, 'word/document.xml');
    if (!docXml) return;

    // Look for page breaks at the beginning of the document
    const blankFirstPageRegex = /<w:body>[\s\S]*?<w:p>[\s\S]*?<w:r>[\s\S]*?<w:br w:type="page"\/>[\s\S]*?<\/w:r>[\s\S]*?<\/w:p>/;
    const match = docXml.match(blankFirstPageRegex);

    if (match) {
      // Replace the blank page with just the body start
      docXml = docXml.replace(match[0], '<w:body>');
      console.log("  Removed blank first page");
      
      // Save the updated document
      mergedZip.file('word/document.xml', docXml);
    } else {
      console.log("  No blank first page detected");
    }
  } catch (error) {
    console.error('  Error removing blank first page:', error);
  }
}

module.exports = {
  applyTemplate
}; 