const dbUrl = "postgresql://neondb_owner:npg_2mJfV9PEOyqk@ep-steep-shadow-az4e9nu5-pooler.c-3.ap-southeast-1.aws.neon.tech/neondb?sslmode=require";

async function queryNeon(sqlQuery, params = []) {
  const directUrl = dbUrl.replace("-pooler", "");
  const parsed = new URL(directUrl.split("?")[0]);
  const host = parsed.hostname;

  const res = await fetch(`https://${host}/sql`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Neon-Connection-String": dbUrl
    },
    body: JSON.stringify({ query: sqlQuery, params })
  });

  if (!res.ok) {
    throw new Error(`Neon error (${res.status}): ${await res.text()}`);
  }
  return await res.json();
}

async function runTest() {
  console.log("Creating table...");
  const tableRes = await queryNeon(`
    CREATE TABLE IF NOT EXISTS invoice_app_data (
      id VARCHAR(50) PRIMARY KEY,
      data JSONB NOT NULL,
      updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
  `);
  console.log("Table created:", tableRes.command);

  console.log("Inserting test payload...");
  const insertRes = await queryNeon(
    `INSERT INTO invoice_app_data (id, data, updated_at)
     VALUES ('test_store', $1, CURRENT_TIMESTAMP)
     ON CONFLICT (id) DO UPDATE
     SET data = EXCLUDED.data, updated_at = CURRENT_TIMESTAMP;`,
    [JSON.stringify({ status: "active", test: true })]
  );
  console.log("Insert result:", insertRes.command);

  console.log("Reading test payload...");
  const selectRes = await queryNeon(`SELECT data FROM invoice_app_data WHERE id = 'test_store';`);
  console.log("Read payload:", selectRes.rows[0]);
}

runTest();
