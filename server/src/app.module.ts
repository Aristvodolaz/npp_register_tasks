import { Module } from '@nestjs/common';
import { ServeStaticModule } from '@nestjs/serve-static';
import * as path from 'path';
import { DbModule } from './db/db.module';
import { OrdersModule } from './orders/orders.module';
import { HealthController } from './health.controller';

@Module({
  imports: [
    DbModule,
    OrdersModule,
    ServeStaticModule.forRoot({
      rootPath: path.resolve(__dirname, '../../client/dist'),
      exclude: ['/api/(.*)'],
    }),
  ],
  controllers: [HealthController],
})
export class AppModule {}
