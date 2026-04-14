[Console]::OutputEncoding = [Text.Encoding]::UTF8
$mojibakeStrings = @(
  '繧｢繧ｯ繧ｻ繧ｹ讓ｩ髯舌′縺ゅｊ縺ｾ縺帙ｸ',
  '譖ｴ譁ｰ縺ｯ5蛻・ｻ･荳翫・髢馴囈繧偵≠縺代※縺上□縺輔＞縲ゅ≠縺ｨ',
  '縺ｧ螳溯｡後〒縺阪∪縺吶・',
  '蛻晄悄繝・・繧ｿ隱ｭ霎ｼ',
  '譖ｴ譁ｰ蠕後ョ繝ｼ繧ｿ遒ｺ隱江',
  'Webhook螳溯｡悟ｾ珪',
  '譖ｴ譁ｰ蜑榊ｾ悟・逅・',
  '譖ｴ譁ｰ髢句ｧ犠',
  '譯井ｻｶ繝・・繧ｿ隱ｭ霎ｼ',
  '譖ｴ譁ｰ蠕梧｡井ｻｶ繝・・繧ｿ遒ｺ隱江'
)
foreach ($s in $mojibakeStrings) {
  $bytes = [Text.Encoding]::GetEncoding('shift_jis').GetBytes($s)
  $result = [Text.Encoding]::UTF8.GetString($bytes)
  Write-Output ($s + ' => ' + $result)
}
