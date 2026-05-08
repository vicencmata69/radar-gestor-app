// Backup automàtic de la base de dades Supabase del Radar-Gestor.
// Es connecta amb les credencials del fitxer .env i descarrega:
//   - licitacions (taula principal amb totes les obres i dates)
//   - tipus
//   - plec_historial (historial d'anàlisis de plecs)
//   - licitacions_log (auditoria de canvis d'estat)
// Genera un fitxer JSON amb timestamp a la carpeta backups/ del projecte.
// Manté un màxim de 30 backups (esborra els més antics).
//
// Ús:
//   node scripts/backup.js
//   o doble clic a auto-backup.bat
//   o programació setmanal amb Windows Task Scheduler

import fs from "node:fs";
import path from "node:path";
import url from "node:url";

const __dirname = path.dirname(url.fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const BACKUP_DIR = path.join(ROOT, "backups");
const MAX_BACKUPS = 30;
const TABLES = ["licitacions", "tipus", "plec_historial", "licitacions_log"];

// Carrega .env
const envPath = path.join(ROOT, ".env");
if (!fs.existsSync(envPath)) {
  console.error("❌ No s'ha trobat .env a:", envPath);
  process.exit(1);
}
const env = Object.fromEntries(
  fs.readFileSync(envPath, "utf8")
    .split(/\r?\n/)
    .filter(l => l.trim() && !l.trim().startsWith("#") && l.includes("="))
    .map(l => {
      const i = l.indexOf("=");
      return [l.slice(0, i).trim(), l.slice(i + 1).trim().replace(/^["']|["']$/g, "")];
    })
);

const SUPABASE_URL = env.VITE_SUPABASE_URL;
const KEY = env.VITE_SUPABASE_ANON_KEY;
if (!SUPABASE_URL || !KEY) {
  console.error("❌ Falten VITE_SUPABASE_URL o VITE_SUPABASE_ANON_KEY a .env");
  process.exit(1);
}

// Crea carpeta de backups si no existeix
if (!fs.existsSync(BACKUP_DIR)) fs.mkdirSync(BACKUP_DIR, { recursive: true });

(async () => {
  const startTime = Date.now();
  const data = { version: 2, date: new Date().toISOString(), source: SUPABASE_URL };
  const counts = {};
  for (const t of TABLES) {
    try {
      const r = await fetch(`${SUPABASE_URL}/rest/v1/${t}?select=*`, {
        headers: { apikey: KEY, Authorization: `Bearer ${KEY}` }
      });
      if (!r.ok) {
        console.warn(`⚠️ Taula ${t}: HTTP ${r.status} — ${await r.text().then(t => t.slice(0, 100))}`);
        data[t] = [];
        counts[t] = 0;
        continue;
      }
      data[t] = await r.json();
      counts[t] = Array.isArray(data[t]) ? data[t].length : 0;
    } catch (e) {
      console.warn(`⚠️ Error a la taula ${t}:`, e.message);
      data[t] = [];
      counts[t] = 0;
    }
  }

  // Compatibilitat amb el format de l'app: licitacions a l'arrel
  data.licitacions = data.licitacions || [];

  const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
  const filename = `backup_servial_auto_${ts}.json`;
  const outPath = path.join(BACKUP_DIR, filename);
  fs.writeFileSync(outPath, JSON.stringify(data, null, 2), "utf8");

  // Rotació: conservem només els MAX_BACKUPS més recents
  const allBackups = fs.readdirSync(BACKUP_DIR)
    .filter(f => f.startsWith("backup_servial_auto_") && f.endsWith(".json"))
    .map(f => ({ name: f, time: fs.statSync(path.join(BACKUP_DIR, f)).mtimeMs }))
    .sort((a, b) => b.time - a.time);
  const toDelete = allBackups.slice(MAX_BACKUPS);
  toDelete.forEach(b => {
    try { fs.unlinkSync(path.join(BACKUP_DIR, b.name)); } catch (e) {}
  });

  const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
  const sizeKB = (fs.statSync(outPath).size / 1024).toFixed(1);
  console.log(`✅ Backup desat: ${filename} (${sizeKB} KB, ${elapsed}s)`);
  console.log(`   Carpeta: ${BACKUP_DIR}`);
  console.log(`   Files per taula:`, counts);
  if (toDelete.length > 0) console.log(`   Eliminats ${toDelete.length} backup(s) antic(s) (>${MAX_BACKUPS}).`);
})().catch(e => {
  console.error("❌ Error fatal:", e);
  process.exit(1);
});
