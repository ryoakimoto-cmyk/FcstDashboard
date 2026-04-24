const fs = require('fs');

let hasError = false;
const files = fs.readdirSync('.').filter(function(file) {
  return /\.(gs|html|md)$/.test(file);
});

files.forEach(function(file) {
  const content = fs.readFileSync(file, 'utf8').replace(/^\ufeff/, '');
  if (content.includes('\ufffd')) {
    console.error('MOJIBAKE DETECTED:', file);
    hasError = true;
    return;
  }
  fs.writeFileSync(file, content, 'utf8');
});

if (hasError) {
  console.error('ENCODING CHECK FAILED.');
  process.exit(1);
}

console.log('OK:', files.length, 'files checked, all UTF-8 clean.');
