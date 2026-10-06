import {
  Body, Controller, Delete, ForbiddenException, Get, HttpCode, Param, ParseIntPipe, Patch, Post,
} from '@nestjs/common';
import { OrdersService } from './orders.service';

@Controller('api/orders')
export class OrdersController {
  constructor(private readonly orders: OrdersService) {}

  @Get()
  list() {
    return this.orders.findAll();
  }

  @Post()
  create(@Body() body: Record<string, unknown>) {
    return this.orders.create(body);
  }

  @Post('bulk')
  @HttpCode(200)
  bulk(@Body() body: { orders: Record<string, unknown>[] }) {
    return this.orders.createMany(Array.isArray(body?.orders) ? body.orders : []);
  }

  @Post('delete-many')
  @HttpCode(200)
  async deleteMany(@Body() body: { ids: number[]; password?: string }) {
    this.checkPassword(body?.password, process.env.DELETE_PASSWORD);
    return { deleted: await this.orders.removeMany(body?.ids ?? []) };
  }

  @Post('clear')
  @HttpCode(200)
  async clear(@Body() body: { password?: string }) {
    this.checkPassword(body?.password, process.env.CLEAR_PASSWORD);
    return { deleted: await this.orders.clear() };
  }

  @Patch(':id')
  update(@Param('id', ParseIntPipe) id: number, @Body() body: Record<string, unknown>) {
    return this.orders.update(id, body);
  }

  @Delete(':id')
  async remove(@Param('id', ParseIntPipe) id: number) {
    return { deleted: await this.orders.removeMany([id]) };
  }

  private checkPassword(given: string | undefined, expected: string | undefined) {
    if (!expected || given !== expected) throw new ForbiddenException('Неверный пароль');
  }
}
