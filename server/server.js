const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3001;

// Enable CORS for frontend requests
app.use(cors());

// Serve static files from the data directory
app.use('/data', express.static(path.join(__dirname, 'data')));

// API endpoint to get the Excel file - dynamically finds any Excel file
app.get('/api/data', (req, res) => {
  const dataDir = path.join(__dirname, 'data');
  
  // Check if data directory exists
  if (!fs.existsSync(dataDir)) {
    return res.status(404).json({ error: 'Data directory not found' });
  }
  
  // Find Excel files in the data directory
  const files = fs.readdirSync(dataDir);
  const excelFiles = files.filter(file => 
    file.endsWith('.xlsx') || file.endsWith('.xls')
  ).filter(file => 
    !file.startsWith('~$') // Exclude temporary Excel files
  );
  
  if (excelFiles.length === 0) {
    return res.status(404).json({ error: 'No Excel files found in data directory' });
  }
  
  // Use the first Excel file found
  const fileName = excelFiles[0];
  const filePath = path.join(dataDir, fileName);
  
  console.log(`Serving Excel file: ${fileName}`);
  
  // Set cache-busting headers to prevent caching
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  res.setHeader('Last-Modified', new Date().toUTCString());
  
  // Send the file
  res.sendFile(filePath);
});

// API endpoint to list available files
app.get('/api/files', (req, res) => {
  const dataDir = path.join(__dirname, 'data');
  
  if (!fs.existsSync(dataDir)) {
    return res.status(404).json({ error: 'Data directory not found' });
  }
  
  const files = fs.readdirSync(dataDir);
  const excelFiles = files.filter(file => 
    file.endsWith('.xlsx') || file.endsWith('.xls')
  ).filter(file => 
    !file.startsWith('~$')
  );
  
  res.json({ files: excelFiles });
});

// API endpoint to get the chunk2 file
app.get('/api/chunk2', (req, res) => {
  const publicDir = path.join(__dirname, '..', 'public');
  const filePath = path.join(publicDir, 'RMT-expense-analysis-chunk2.xlsx');
  
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'Chunk2 file not found' });
  }
  
  console.log(`Serving chunk2 file: RMT-expense-analysis-chunk2.xlsx`);
  
  // Set cache-busting headers to prevent caching
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  res.setHeader('Last-Modified', new Date().toUTCString());
  
  // Send the file
  res.sendFile(filePath);
});

// API endpoint to get the Financial Performance Data file
app.get('/api/financial-performance', (req, res) => {
  const publicDir = path.join(__dirname, '..', 'public');
  const filePath = path.join(publicDir, 'Financial Performance Data.xlsx');
  
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'Financial Performance Data file not found' });
  }
  
  console.log(`Serving Financial Performance Data file: Financial Performance Data.xlsx`);
  
  // Set cache-busting headers to prevent caching
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  res.setHeader('Last-Modified', new Date().toUTCString());
  
  // Send the file
  res.sendFile(filePath);
});

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ status: 'OK', message: 'Backend server is running' });
});

app.listen(PORT, () => {
  console.log(`Backend server running on port ${PORT}`);
  console.log(`Excel file available at: http://localhost:${PORT}/api/data`);
}); 