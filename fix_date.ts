import fs from 'fs';

function fixDatesInFile(filePath) {
   let content = fs.readFileSync(filePath, 'utf8');
   content = content.replace(/p\.date\.toLocaleDateString\('es-ES',\s*\{\s*month:\s*'short',\s*year:\s*'numeric'\s*\}\)/g, 'formatDateKey(p.date)');
   content = content.replace(/point\.date\.toLocaleDateString\('es-ES',\s*\{\s*month:\s*'short',\s*year:\s*'numeric'\s*\}\)/g, 'formatDateKey(point.date)');
   content = content.replace(/dateObj\.toLocaleDateString\('es-ES',\s*\{\s*month:\s*'short',\s*year:\s*'numeric'\s*\}\)/g, 'formatDateKey(dateObj)');
   content = content.replace(/d\.toLocaleDateString\('es-ES',\s*\{\s*month:\s*'short',\s*year:\s*'numeric'\s*\}\)/g, 'formatDateKey(d)');
   content = content.replace(/curDate\.toLocaleDateString\('es-ES',\s*\{\s*month:\s*'short',\s*year:\s*'numeric'\s*\}\)/g, 'formatDateKey(curDate)');
   content = content.replace(/date\.toLocaleDateString\('es-ES',\s*\{\s*month:\s*'short',\s*year:\s*'numeric'\s*\}\)/g, 'formatDateKey(date)');
   content = content.replace(/dSort\.toLocaleDateString\('es-ES',\s*\{\s*month:\s*'short',\s*year:\s*'numeric'\s*\}\)/g, 'formatDateKey(dSort)');
   content = content.replace(/item\.sortDate\.toLocaleDateString\('es-ES',\s*\{\s*month:\s*'short',\s*year:\s*'numeric'\s*\}\)/g, 'formatDateKey(item.sortDate)');
   
   fs.writeFileSync(filePath, content);
}

fixDatesInFile('financialEngine.js');
fixDatesInFile('main.js');
console.log('Done replacement!');
