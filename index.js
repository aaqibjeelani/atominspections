const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { mergeDocuments, cleanupFiles } = require('./wordMerger');

const app = express();
const port = 3000;

// Create uploads directory if it doesn't exist
if (!fs.existsSync('./uploads')) {
  fs.mkdirSync('./uploads');
}

// Set up storage for uploaded files
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, './uploads');
  },
  filename: (req, file, cb) => {
    // Create a unique filename with original extension
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, uniqueSuffix + path.extname(file.originalname));
  }
});

// Create the upload instance
const upload = multer({ 
  storage: storage,
  fileFilter: (req, file, cb) => {
    // Only accept Word documents
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
      cb(null, true);
    } else {
      cb(new Error('Only .docx format allowed!'), false);
    }
  }
});

// Serve static files
app.use(express.static('public'));

// Handle file uploads and merging
app.post('/merge', upload.array('documents', 10), async (req, res) => {
  try {
    if (!req.files || req.files.length < 1) {
      return res.status(400).send('Please upload at least one document');
    }

    const templatePath = path.join(__dirname, 'sample', 'template.docx');
    const outputPath = path.join(__dirname, 'uploads', 'merged-document.docx');
    
    // Get paths of uploaded files
    const documentPaths = req.files.map(file => file.path);
    
    // Merge the documents using the utility
    await mergeDocuments(templatePath, documentPaths, outputPath);
    
    // Send the merged file as a download
    res.download(outputPath, 'merged-document.docx', (err) => {
      if (err) {
        console.error('Error downloading file:', err);
        return res.status(500).send('Error downloading the merged file');
      }
      
      // Clean up uploaded files after successful download
      cleanupFiles(documentPaths);
      
      // Also clean up the output file after some delay
      setTimeout(() => {
        cleanupFiles([outputPath]);
      }, 30000); // Clean after 30 seconds
    });
  } catch (error) {
    console.error('Error merging documents:', error);
    return res.status(500).send(`Error merging documents: ${error.message}`);
  }
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).send(`An error occurred: ${err.message}`);
});

// Create a simple HTML frontend
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Start the server
app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
  console.log(`Using template from: ${path.join(__dirname, 'sample', 'template.docx')}`);
}); 