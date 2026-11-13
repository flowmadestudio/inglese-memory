-- Already applied to the D1 database "inglese-memory" (13 Nov 2026).
CREATE TABLE IF NOT EXISTS users (uid TEXT PRIMARY KEY, email TEXT, created_at INTEGER NOT NULL, last_seen INTEGER);
CREATE TABLE IF NOT EXISTS progress (uid TEXT NOT NULL, word TEXT NOT NULL, data TEXT NOT NULL, updated_at INTEGER NOT NULL, PRIMARY KEY (uid, word));
CREATE TABLE IF NOT EXISTS queue (uid TEXT NOT NULL, position INTEGER NOT NULL, word TEXT NOT NULL, PRIMARY KEY (uid, word));
