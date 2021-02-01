const fs = require('fs');
const https = require('https');
const express = require('express');
const bodyParser =require('body-parser');
const cors = require('cors');
const path = require('path');

const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());
// app.use(morgan('dev'));
app.use(express.static('assets'));
app.use(express.static('dist'));
/* Turn off caching when debugging */
app.use(function (req, res, next) {
    res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    res.header('Expires', '-1');
    res.header('Pragma', 'no-cache');
    next()
});

const cert = {
//    key: fs.readFileSync(path.resolve('cert/key.pem')),
    key: fs.readFileSync(path.resolve('cert/server.key')),
//    cert: fs.readFileSync(path.resolve('cert/cert.pem'))
    cert: fs.readFileSync(path.resolve('cert/server.crt'))
};
https.createServer(cert, app).listen(3000, () => console.log('Server running on 3000'));
