const express = require('express');
const path = require('path');
const app = express();
const port = 3000; // You can change the port as needed
const data = require('../assets/data.js'); // Adjust the path to your data.js file

console.log(__dirname)

// Serve static files from the 'public' directory
//app.use(express.static(path.join(__dirname, 'public')));

// Serve static files from the 'dist' directory if using Webpack
app.use(express.static(path.join(__dirname, 'dist')));

// // Serve taskpane.html at the root
// app.get('/', (req, res) => {
//   res.sendFile(path.join(__dirname, '..', 'client', 'taskpane', 'taskpane.html'));
// });

app.get('/api/data', (req, res) => {
  res.json(data);
});

app.listen(port, () => {
  console.log(`Server running on https://localhost:${port}`);
});