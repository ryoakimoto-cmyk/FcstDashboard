const fs = require('fs');
const EXPECTED = [
  'FCST調整', 'Export待機', '集計キャッシュ', '部門マスタ',
  '月次目標マスタ', '変更ログ', '目標',
  'FCSTスナップショット', '案件リストスナップショット',
  'Export待機_提案商品', 'SFユーザー',
  'SFデータ更新_BO', 'SFデータ更新_SS', 'SFデータ更新_SSCS', 'SFデータ更新_CO'
];
let hasError = false;
const files = fs.readdirSync('.').filter(function(file) {
  return /\.(gs|html)$/.test(file);
});

files.forEach(function(file) {
  let content = fs.readFileSync(file, 'utf8').replace(/^\ufeff/, '');
  if (content.includes('\ufffd')) {
    console.error('MOJIBAKE DETECTED:', file);
    hasError = true;
    return;
  }
  fs.writeFileSync(file, content, 'utf8');
});

const all = files.map(function(file) {
  return fs.readFileSync(file, 'utf8');
}).join('');

EXPECTED.forEach(function(text) {
  if (!all.includes(text)) {
    console.error('MISSING STRING:', text);
    hasError = true;
  }
});

if (hasError) {
  console.error('ENCODING CHECK FAILED.');
  process.exit(1);
} else {
  console.log('OK:', files.length, 'files checked, all UTF-8 clean.');
}
