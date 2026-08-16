const fs = require('fs');
let htmlText = fs.readFileSync('index.html', 'utf8');

// 1. Change #auth-page.active to use flex-direction: column and remove center alignment
htmlText = htmlText.replace(
  /#auth-page\.active\{display:flex;min-height:100vh;align-items:center;justify-content:center;padding:24px;background:var\(--bg\);padding-top:calc\(var\(--nav-h\)\+24px\)\}/,
  '#auth-page.active{display:flex;flex-direction:column;min-height:100vh;padding:24px;background:var(--bg);padding-top:calc(var(--nav-h)+24px)}'
);

// 2. Change .auth-wrap to use margin: auto to perfectly center vertically but respect top padding on small screens
htmlText = htmlText.replace(
  /\.auth-wrap\{width:100%;max-width:480px;background:var\(--surface\);border:1px solid var\(--border\);border-radius:20px;padding:36px 28px;box-shadow:var\(--sh-lg\);position:relative;overflow:hidden;margin-top:24px;\}/,
  '.auth-wrap{width:100%;max-width:480px;background:var(--surface);border:1px solid var(--border);border-radius:20px;padding:36px 28px;box-shadow:var(--sh-lg);position:relative;overflow:hidden;margin:auto;}'
);

fs.writeFileSync('index.html', htmlText, 'utf8');
console.log('Fixed Auth Page Overlap.');
