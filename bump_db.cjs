const fs = require('fs');

function bumpVersion(file) {
    let content = fs.readFileSync(file, 'utf8');
    content = content.replace(/PlanetaAzulDB',\s*[0-4]/g, "PlanetaAzulDB', 5");
    fs.writeFileSync(file, content);
}

bumpVersion('main.js');
bumpVersion('resumenComercialEngine.js');
console.log("Bumped IndexedDB version to 5");
