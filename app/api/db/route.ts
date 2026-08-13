import { NextResponse } from "next/server";

async function queryNeon(connectionString: string, sqlQuery: string, params: unknown[] = []) {
  const cleanConnStr = connectionString.replace("-pooler", "");
  const parsed = new URL(cleanConnStr.split("?")[0]);
  const host = parsed.hostname;

  const response = await fetch(`https://${host}/sql`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Neon-Connection-String": cleanConnStr
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
  const dbUrl = process.env.DATABASE_URL;
  if (!dbUrl) {
    return NextResponse.json({
      connected: false,
      mode: "local",
      message: "DATABASE_URL is not set. Running in browser LocalStorage mode."
    });
  }

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
          // keep string or null
        }
      }
    }

    return NextResponse.json({
      connected: true,
      mode: "neon",
      data: storedData,
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
  const dbUrl = process.env.DATABASE_URL;
  if (!dbUrl) {
    return NextResponse.json({
      success: false,
      mode: "local",
      message: "DATABASE_URL is not configured on Vercel/env. Operating in LocalStorage mode."
    });
  }

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
