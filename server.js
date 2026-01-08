const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 3000;
const ROOT_DIR = __dirname;

// MIME 타입 매핑
const MIME_TYPES = {
  '.html': 'text/html',
  '.css': 'text/css',
  '.js': 'application/javascript',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.csv': 'text/csv',
  '.ico': 'image/x-icon'
};

const server = http.createServer((req, res) => {
  // URL 파싱 (쿼리스트링 제거)
  let filePath = req.url.split('?')[0];
  
  // 루트 경로는 index.html로
  if (filePath === '/') {
    filePath = '/index.html';
  }

  // 실제 파일 경로 생성
  const fullPath = path.join(ROOT_DIR, filePath);
  
  // 보안: 상위 디렉토리 접근 방지
  if (!fullPath.startsWith(ROOT_DIR)) {
    res.writeHead(403, { 'Content-Type': 'text/plain' });
    res.end('Forbidden');
    return;
  }

  // 파일 읽기
  fs.readFile(fullPath, (err, data) => {
    if (err) {
      // 파일이 없으면 index.html로 폴백 (SPA 라우팅)
      if (err.code === 'ENOENT') {
        const indexPath = path.join(ROOT_DIR, 'index.html');
        fs.readFile(indexPath, (err, data) => {
          if (err) {
            res.writeHead(404, { 'Content-Type': 'text/plain' });
            res.end('404 Not Found');
          } else {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(data);
          }
        });
      } else {
        res.writeHead(500, { 'Content-Type': 'text/plain' });
        res.end('500 Internal Server Error');
      }
      return;
    }

    // MIME 타입 결정
    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    res.writeHead(200, { 'Content-Type': contentType });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`🚀 정적 서버가 http://localhost:${PORT} 에서 실행 중입니다.`);
  console.log(`📁 서비스 디렉토리: ${ROOT_DIR}`);
});





