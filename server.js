const express = require('express');
const path = require('path');

const app = express();
const PORT = 3000;

// 정적 파일 제공 (HTML, CSS, JS, CSV 등)
app.use(express.static(__dirname));

// 모든 경로를 index.html로 라우팅 (SPA 지원)
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, () => {
  console.log(`🚀 서버가 http://localhost:${PORT} 에서 실행 중입니다.`);
  console.log(`📁 정적 파일 제공: ${__dirname}`);
});



