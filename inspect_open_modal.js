const fs = require('fs');
const html = fs.readFileSync('index.html', 'utf8');
const lines = html.split('\n');
lines.forEach((l, i) => {
  if (l.includes('function openModal(')) {
    console.log(lines.slice(i, i + 25).join('\n'));
  }
});
