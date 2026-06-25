// Minimal Express backend for IMPACT Tracker.
// Serves the built Vite frontend (../dist) and a tiny JSON state API backed by Postgres.
import express from 'express';
import pg from 'pg';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const distDir = path.join(__dirname, '..', 'dist');

const { Pool } = pg;
const connectionString = process.env.DATABASE_URL;
// Render's external host needs SSL; the internal host (same region) does not.
const useSsl = /render\.com/.test(connectionString || '');
const pool = connectionString
  ? new Pool({ connectionString, ssl: useSsl ? { rejectUnauthorized: false } : false })
  : null;

// State keys persisted as one JSONB document each.
const KEYS = ['tasks', 'subtypes', 'members'];

async function initDb() {
  if (!pool) {
    console.warn('[server] DATABASE_URL not set — state API will return empty data.');
    return;
  }
  await pool.query(`
    CREATE TABLE IF NOT EXISTS app_state (
      key        TEXT PRIMARY KEY,
      value      JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);
  console.log('[server] database ready');
}

const app = express();
app.use(express.json({ limit: '5mb' }));

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, db: Boolean(pool) });
});

// Return the full shared state: { tasks, subtypes, members }
app.get('/api/state', async (_req, res) => {
  if (!pool) return res.json({});
  try {
    const { rows } = await pool.query('SELECT key, value FROM app_state WHERE key = ANY($1)', [KEYS]);
    const out = {};
    for (const r of rows) out[r.key] = r.value;
    res.json(out);
  } catch (err) {
    console.error('[server] GET /api/state failed:', err.message);
    res.status(500).json({ error: 'read_failed' });
  }
});

// Upsert any subset of { tasks, subtypes, members }
app.put('/api/state', async (req, res) => {
  if (!pool) return res.status(503).json({ error: 'no_database' });
  const body = req.body || {};
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    for (const key of KEYS) {
      if (body[key] === undefined) continue;
      await client.query(
        `INSERT INTO app_state (key, value, updated_at) VALUES ($1, $2, now())
         ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = now()`,
        [key, JSON.stringify(body[key])]
      );
    }
    await client.query('COMMIT');
    res.json({ ok: true });
  } catch (err) {
    await client.query('ROLLBACK').catch(() => {});
    console.error('[server] PUT /api/state failed:', err.message);
    res.status(500).json({ error: 'write_failed' });
  } finally {
    client.release();
  }
});

// Serve the built frontend.
app.use(express.static(distDir));
// SPA fallback: any non-API GET returns index.html.
app.use((req, res, next) => {
  if (req.method !== 'GET' || req.path.startsWith('/api')) return next();
  res.sendFile(path.join(distDir, 'index.html'));
});

const port = process.env.PORT || 3000;
initDb()
  .catch((err) => console.error('[server] initDb error:', err.message))
  .finally(() => {
    app.listen(port, () => console.log(`[server] listening on :${port}`));
  });
