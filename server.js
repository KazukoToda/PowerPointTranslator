require('dotenv').config();
const express = require('express');
const fileUpload = require('express-fileupload');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const { translatePowerPoint } = require('./translator');

const app = express();
const PORT = process.env.PORT || 3000;

// Create uploads directory if it doesn't exist
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(fileUpload({
  createParentPath: true,
  limits: { 
    fileSize: 50 * 1024 * 1024 // 50MB max file size
  },
}));
app.use(express.static('public'));

// Routes
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Handle PowerPoint upload and translation
app.post('/translate', async (req, res) => {
  try {
    if (!req.files || Object.keys(req.files).length === 0) {
      return res.status(400).json({ error: 'No file was uploaded.' });
    }

    const powerPointFile = req.files.file;
    
    // Save the uploaded file
    const uploadPath = path.join(uploadsDir, powerPointFile.name);
    await powerPointFile.mv(uploadPath);
    
    // Translate the PowerPoint
    const translatedFilePath = await translatePowerPoint(uploadPath);
    
    // Return the path to the translated file
    return res.json({ 
      success: true, 
      filePath: path.basename(translatedFilePath),
      message: 'PowerPoint file translated successfully.' 
    });
  } catch (err) {
    console.error('Error processing file:', err);
    return res.status(500).json({ 
      error: 'Failed to process the PowerPoint file.', 
      details: err.message 
    });
  }
});

// Route for downloading the translated file
app.get('/download/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(uploadsDir, filename);
  
  if (fs.existsSync(filePath)) {
    res.download(filePath);
  } else {
    res.status(404).json({ error: 'File not found.' });
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});