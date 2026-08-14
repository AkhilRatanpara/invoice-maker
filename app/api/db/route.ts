import { NextResponse } from "next/server";

const DEFAULT_DATABASE_URL =
  "postgresql://neondb_owner:npg_2mJfV9PEOyqk@ep-steep-shadow-az4e9nu5-pooler.c-3.ap-southeast-1.aws.neon.tech/neondb?sslmode=require";

const getDbUrl = () => process.env.DATABASE_URL || DEFAULT_DATABASE_URL;

async function queryNeon(connectionString: string, sqlQuery: string, params: unknown[] = []) {
  const directUrl = connectionString.replace("-pooler", "");
  const parsed = new URL(directUrl.split("?")[0]);
  const host = parsed.hostname;

  const response = await fetch(`https://${host}/sql`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Neon-Connection-String": connectionString
    },
    body: JSON.stringify({ query: sqlQuery, params })
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Neon DB error (${response.status}): ${errorText}`);
  }

  return await response.json();
}

export async function GET() {
  const dbUrl = getDbUrl();

  try {
    await queryNeon(
      dbUrl,
      `CREATE TABLE IF NOT EXISTS invoice_app_data (
        id VARCHAR(50) PRIMARY KEY,
        data JSONB NOT NULL,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );`
    );

    const result = await queryNeon(
      dbUrl,
      `SELECT data FROM invoice_app_data WHERE id = 'main_store' LIMIT 1;`
    );

    let storedData = null;
    if (result && Array.isArray(result.rows) && result.rows.length > 0) {
      storedData = result.rows[0].data || result.rows[0][0] || null;
      if (typeof storedData === "string") {
        try {
          storedData = JSON.parse(storedData);
        } catch {
          // keep string
        }
      }
    }

    let storageInfo = { prettySize: "0 B", bytes: 0 };
    try {
      const sizeResult = await queryNeon(
        dbUrl,
        `SELECT pg_size_pretty(pg_total_relation_size('invoice_app_data')) as pretty_size, pg_total_relation_size('invoice_app_data') as bytes;`
      );
      if (sizeResult && Array.isArray(sizeResult.rows) && sizeResult.rows.length > 0) {
        const row = sizeResult.rows[0];
        storageInfo = {
          prettySize: String(row.pretty_size ?? row[0] ?? "0 B"),
          bytes: Number(row.bytes ?? row[1] ?? 0)
        };
      }
    } catch {
      // Ignore size query fallback
    }

    return NextResponse.json({
      connected: true,
      mode: "neon",
      data: storedData,
      storageInfo,
      message: "Successfully connected to Neon Postgres database!"
    });
  } catch (error) {
    return NextResponse.json({
      connected: false,
      mode: "local",
      error: error instanceof Error ? error.message : "Failed to connect to Neon database"
    });
  }
}

export async function POST(request: Request) {
  const dbUrl = getDbUrl();

  try {
    const body = await request.json();

    await queryNeon(
      dbUrl,
      `CREATE TABLE IF NOT EXISTS invoice_app_data (
        id VARCHAR(50) PRIMARY KEY,
        data JSONB NOT NULL,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );`
    );

    await queryNeon(
      dbUrl,
      `INSERT INTO invoice_app_data (id, data, updated_at)
       VALUES ('main_store', $1, CURRENT_TIMESTAMP)
       ON CONFLICT (id) DO UPDATE
       SET data = EXCLUDED.data, updated_at = CURRENT_TIMESTAMP;`,
      [JSON.stringify(body)]
    );

    return NextResponse.json({
      success: true,
      mode: "neon",
      message: "Data successfully synced with Neon Postgres database!"
    });
  } catch (error) {
    return NextResponse.json({
      success: false,
      error: error instanceof Error ? error.message : "Database sync error"
    });
  }
}
