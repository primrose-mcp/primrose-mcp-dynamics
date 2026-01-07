/**
 * Sales Order Tools
 *
 * MCP tools for sales order management in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

export function registerOrderTools(server: McpServer, client: CrmClient): void {
  server.tool(
    'dynamics_list_orders',
    'List sales orders from Dynamics 365 with pagination.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listOrders({ limit, cursor });
        return formatResponse(result, format, 'orders');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_order',
    'Get a single sales order by ID.',
    { id: z.string(), format: z.enum(['json', 'markdown']).default('json') },
    async ({ id, format }) => {
      try {
        const order = await client.getOrder(id);
        return formatResponse(order, format, 'order');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_order',
    'Create a new sales order in Dynamics 365.',
    {
      name: z.string().describe('Order name'),
      description: z.string().optional(),
      quoteId: z.string().optional(),
      opportunityId: z.string().optional(),
      customerAccountId: z.string().optional(),
      priceLevelId: z.string().optional(),
      requestDeliveryBy: z.string().optional(),
    },
    async (input) => {
      try {
        const order = await client.createOrder(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Order created', order }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_update_order',
    'Update an existing sales order.',
    {
      id: z.string(),
      name: z.string().optional(),
      description: z.string().optional(),
      requestDeliveryBy: z.string().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const order = await client.updateOrder(id, input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Order updated', order }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_delete_order',
    'Delete a sales order from Dynamics 365.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.deleteOrder(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: `Order ${id} deleted` }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_list_order_details',
    'List line items for a sales order.',
    { orderId: z.string() },
    async ({ orderId }) => {
      try {
        const details = await client.listOrderDetails(orderId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, count: details.length, details }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_add_order_detail',
    'Add a line item to a sales order.',
    {
      salesOrderId: z.string(),
      quantity: z.number(),
      productId: z.string().optional(),
      productDescription: z.string().optional(),
      pricePerUnit: z.number().optional(),
      manualDiscountAmount: z.number().optional(),
      tax: z.number().optional(),
    },
    async (input) => {
      try {
        const detail = await client.addOrderDetail(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Order detail added', detail }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_fulfill_order',
    'Mark a sales order as fulfilled.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.fulfillOrder(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Order fulfilled' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_cancel_order',
    'Cancel a sales order.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.cancelOrder(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Order cancelled' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_convert_order_to_invoice',
    'Convert a sales order to an invoice.',
    { orderId: z.string() },
    async ({ orderId }) => {
      try {
        const result = await client.convertOrderToInvoice(orderId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Order converted to invoice', ...result }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
