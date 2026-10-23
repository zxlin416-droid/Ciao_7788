/**
 * 打包成免安裝 Node.js 的發佈資料夾：dist/執課表彙整/（含 exe、範本、報表資料夾、啟動檔）
 * 用法：npm run build
 */
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const dist = path.join(__dirname, 'dist');
const pkgDir = path.join(dist, '執課表彙整');
const run = (c) => execSync(c, { stdio: 'inherit', cwd: __dirname });

fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(path.join(pkgDir, '報表'), { recursive: true });

run('npx esbuild 彙整執課表.js --bundle --platform=node --target=node22 --outfile=dist/bundle.js --log-level=warning');
run(`npx pkg dist/bundle.js --targets node22-win-x64 --output "${path.join(pkgDir, 'ClassReport.exe')}"`);

fs.copyFileSync(path.join(__dirname, '執課表範本.xlsx'), path.join(pkgDir, '執課表範本.xlsx'));
fs.writeFileSync(path.join(pkgDir, '報表', '請把兩份報表放這裡.txt'),
  '把「私人課表預約明細拆帳版」與「團體課預約報表」兩份 xlsx 放在這個資料夾，然後雙擊上一層的「執行彙整.bat」。\r\n');
fs.writeFileSync(path.join(pkgDir, '執行彙整.bat'),
  '@echo off\r\nchcp 65001 >nul\r\ncd /d "%~dp0"\r\n.\\ClassReport.exe\r\necho.\r\npause\r\n');
fs.rmSync(path.join(dist, 'bundle.js'));
console.log('\n完成：' + pkgDir);
