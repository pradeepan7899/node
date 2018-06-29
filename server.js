const http = require('http');

const app  = require('./app');

const port = "http://ec2-52-221-204-47.ap-southeast-1.compute.amazonaws.com:3000";

const server = http.createServer(app);

server.listen(port);