import { Injectable, Logger, OnModuleDestroy, OnModuleInit } from '@nestjs/common';
import * as sql from 'mssql';

const bool = (v: string | undefined, def: boolean) =>
  v === undefined || v === '' ? def : v.toLowerCase() === 'true';

@Injectable()
export class DbService implements OnModuleInit, OnModuleDestroy {
  private readonly log = new Logger('DB');
  pool!: sql.ConnectionPool;

  async onModuleInit() {
    const e = process.env;
    this.pool = new sql.ConnectionPool({
      server: e.DB_SERVER!,
      port: Number(e.DB_PORT) || 1433,
      database: e.DB_NAME || e.DB_DATABASE,
      user: e.DB_USER,
      password: e.DB_PASSWORD,
      pool: {
        max: Number(e.DB_POOL_MAX) || 10,
        min: Number(e.DB_POOL_MIN) || 0,
        idleTimeoutMillis: Number(e.DB_IDLE_TIMEOUT) || 30000,
      },
      options: {
        encrypt: bool(e.DB_ENCRYPT, true),
        trustServerCertificate: bool(e.DB_TRUST_SERVER_CERTIFICATE, true),
        enableArithAbort: bool(e.DB_ENABLE_ARITH_ABORT, true),
      },
    });
    this.pool.on('error', (err) => this.log.error(err.message));
    await this.pool.connect();
    this.log.log(`Подключено: ${e.DB_SERVER}:${e.DB_PORT}/${e.DB_NAME || e.DB_DATABASE}`);
  }

  async onModuleDestroy() {
    await this.pool?.close();
  }

  request() {
    return this.pool.request();
  }

  async query<T = any>(text: string): Promise<T[]> {
    const r = await this.pool.request().query(text);
    return r.recordset as T[];
  }
}
