import 'reflect-metadata';
import * as path from 'path';
import * as dotenv from 'dotenv';
// .env лежит в корне проекта (на уровень выше server/)
dotenv.config({ path: path.resolve(__dirname, '../../.env') });

import { NestFactory } from '@nestjs/core';
import { json } from 'express';
import { AppModule } from './app.module';

async function bootstrap() {
  const app = await NestFactory.create(AppModule);
  app.use(json({ limit: '20mb' })); // импорт больших Excel-файлов
  app.enableShutdownHooks();
  const port = Number(process.env.PORT) || 3043;
  await app.listen(port, '0.0.0.0');
  console.log(`НПП Регистр заявок: http://0.0.0.0:${port}`);
}
bootstrap();
