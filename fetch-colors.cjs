const https = require('https');

https.get('https://help4u.com.br/', (res) => {
  let data = '';
  res.on('data', (chunk) => {
    data += chunk;
  });
  res.on('end', () => {
    const rawMatches = data.match(/#[A-Fa-f0-9]{6}/g) || [];
    
    // Count occurrences
    const counts = {};
    for (const c of rawMatches) {
        let code = c.toLowerCase();
        counts[code] = (counts[code] || 0) + 1;
    }
    
    const sorted = Object.entries(counts).sort((a,b) => b[1] - a[1]);
    console.log('Top hex colors:', sorted.slice(0, 10));
  });
}).on('error', (err) => {
  console.log('Error: ' + err.message);
});
