// Builds the single-file viewer: viewer-template.html + inlined reader -> dist/index.html
import fs from 'fs';
const tables = fs.readFileSync('tables.js', 'utf8');
const reader = fs.readFileSync('mpp-reader.js', 'utf8').replace('/* @@TABLES@@ */', tables);
fs.writeFileSync('dist-reader.js', reader);
let html = fs.readFileSync('viewer-template.html', 'utf8');
const swap = (a, b) => { if (!html.includes(a)) throw new Error('missing: ' + a.slice(0, 60)); html = html.replace(a, b); };
swap(`<p>See the task list, timeline and dependencies without opening Project. The file is read once and isn't kept.</p>`,
     `<p>See the task list, timeline and dependencies without opening Project. Your file is opened in this browser and is never uploaded anywhere.</p>`);
swap(`<div class="types">.mpp from Project 98 to 2021 and Project Online, plus Project XML, MPX and Primavera XER</div>`,
     `<div class="types">.mpp files from Project 2010 or newer (including Microsoft 365), or Project XML (.xml) from any version</div>`);
swap(`accept=".mpp,.mpt,.xml,.mpx,.xer,.pmxml,.pp,.gan,.planner,.pod,.cdpx,.sdef"`, `accept=".mpp,.mpt,.xml"`);
swap(`    const fd = new FormData(); fd.append('file', file);
    const r = await fetch('api/parse', { method:'POST', body:fd });
    let j; try{ j = await r.json(); }catch{ throw new Error(\`The server returned an unexpected response (\${r.status}).\`); }
    if(!r.ok) throw new Error(j.error || \`Upload failed (\${r.status}).\`);
    load(j);`,
`    if(file.size > 200*1024*1024) throw new Error('That file is larger than 200 MB.');
    await new Promise(r=>setTimeout(r,30));          // let the "Reading…" state paint
    const j = await MPPReader.read(file);
    if(!j.tasks.some(t=>!t.blank && t.start)) throw new Error('This file has no scheduled tasks to show.');
    load(j);`);
swap(`$('#err').textContent = e.message === 'Failed to fetch' ? "The server can't be reached. Check that it's running and try again." : e.message;`,
     `$('#err').textContent = e.message;`);
// untitled tasks
swap(`<span class="lbl" title="\${esc(t.name)}">\${hl(t.name||'')}</span>`, `<span class="lbl" title="\${esc(t.name||'')}">\${t.name?hl(t.name):'<i style="color:var(--ink-3)">Untitled task</i>'}</span>`);
swap(`<script>\nconst $ =`, `<script>\n${reader}\n</script>\n<script>\nconst $ =`);
fs.mkdirSync('dist', { recursive: true });
fs.writeFileSync('dist/index.html', html);
console.log('dist/index.html', (html.length/1024).toFixed(0)+' KB');
