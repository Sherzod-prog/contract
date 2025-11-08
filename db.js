const pg = require("pg");
const dotenv = require("dotenv");
dotenv.config();

const pool = new pg.Pool({
  connectionString: process.env.DATABASE_URL,
});
pool.query(`CREATE TABLE IF NOT EXISTS contracts (
    id SERIAL PRIMARY KEY,
    date VARCHAR(50),
    bank_name VARCHAR(100),
    contact VARCHAR(100)
)`);
module.exports = { pool };
