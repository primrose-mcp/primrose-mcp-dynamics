/**
 * Invoice Tools
 *
 * MCP tools for invoice management in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

export function registerInvoiceTools(server: McpServer, client: CrmClient): void {
  server.tool(
    'dynamics_list_invoices',
    'List invoices from Dynamics 365 with pagination.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listInvoices({ limit, cursor });
        return formatResponse(result, format, 'invoices');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_invoice',
    'Get a single invoice by ID.',
    { id: z.string(), format: z.enum(['json', 'markdown']).default('json') },
    async ({ id, format }) => {
      try {
        const invoice = await client.getInvoice(id);
        return formatResponse(invoice, format, 'invoice');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_invoice',
    'Create a new invoice in Dynamics 365.',
    {
      name: z.string().describe('Invoice name'),
      description: z.string().optional(),
      salesOrderId: z.string().optional(),
      customerAccountId: z.string().optional(),
      priceLevelId: z.string().optional(),
      dueDate: z.string().optional(),
    },
    async (input) => {
      try {
        const invoice = await client.createInvoice(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Invoice created', invoice }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_update_invoice',
    'Update an existing invoice.',
    {
      id: z.string(),
      name: z.string().optional(),
      description: z.string().optional(),
      dueDate: z.string().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const invoice = await client.updateInvoice(id, input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Invoice updated', invoice }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_delete_invoice',
    'Delete an invoice from Dynamics 365.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.deleteInvoice(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: `Invoice ${id} deleted` }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_lock_invoice_pricing',
    'Lock the pricing on an invoice.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.lockInvoicePricing(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Invoice pricing locked' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_cancel_invoice',
    'Cancel an invoice.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.cancelInvoice(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Invoice cancelled' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
