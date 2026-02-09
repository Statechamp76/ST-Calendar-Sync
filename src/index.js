const http = require('http');
const app = require('./app');
const { loadConfig } = require('./config');

const config = loadConfig();
const PORT = config.port || 8080;

const server = http.createServer((req, res) => {
  if (req.url === '/health') {
    res.writeHead(200);
    return res.end('ok');
  }

  return app(req, res);
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Server listening on port ${PORT}`);
});
