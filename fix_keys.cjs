const fs = require('fs');
let s = fs.readFileSync('main.js', 'utf8');

// 1. Where it saves to indexeddb:
s = s.replace('MASTER_STANDALONE_CXP_DATA', 'CXP_STANDALONE_KEY');

// 2. Where it gets read from indexeddb (lines 1342+):
// Replace: window.cachedStandaloneCxp = cxpCachedRecord.data;
// With:    window.cxpStandaloneData = cxpCachedRecord.data; window.hasCxpAccess = true;
s = s.replace(
    'window.cachedStandaloneCxp = cxpCachedRecord.data;', 
    'window.cxpStandaloneData = cxpCachedRecord.data; window.hasCxpAccess = true;'
);

fs.writeFileSync('main.js', s);
