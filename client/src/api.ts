import type { Order, OrderInput } from './utils';

async function req<T>(url: string, method = 'GET', body?: unknown): Promise<T> {
  const res = await fetch(url, {
    method,
    headers: body ? { 'Content-Type': 'application/json' } : undefined,
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    let msg = res.statusText;
    try { msg = (await res.json()).message || msg; } catch { /* не JSON */ }
    throw new Error(Array.isArray(msg) ? msg.join(', ') : msg);
  }
  return res.json() as Promise<T>;
}

export const api = {
  list: () => req<Order[]>('/api/orders'),
  create: (o: OrderInput) => req<Order>('/api/orders', 'POST', o),
  update: (id: number, o: OrderInput) => req<Order>(`/api/orders/${id}`, 'PATCH', o),
  remove: (id: number) => req<{ deleted: number }>(`/api/orders/${id}`, 'DELETE'),
  bulk: (orders: OrderInput[]) => req<{ imported: number; skipped: number }>('/api/orders/bulk', 'POST', { orders }),
  deleteMany: (ids: number[], password: string) =>
    req<{ deleted: number }>('/api/orders/delete-many', 'POST', { ids, password }),
  clear: (password: string) => req<{ deleted: number }>('/api/orders/clear', 'POST', { password }),
};
