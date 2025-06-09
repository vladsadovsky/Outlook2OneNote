const fs = require('fs');
const https = require('https');
const express = require('express');
const app = express();

app.use(express.static('public'));

https.createServer({
  key: fs.readFileSync('certs/localhost.key'),
  cert: fs.readFileSync('certs/localhost.crt')
}, app).listen(3000, () => {
  console.log('✅ HTTPS Server running at https://localhost:3000');
});
