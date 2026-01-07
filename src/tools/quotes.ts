/**
 * Quote Tools
 *
 * MCP tools for quote management in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

export function registerQuoteTools(server: McpServer, client: CrmClient): void {
  server.tool(
    'dynamics_list_quotes',
    'List quotes from Dynamics 365 with pagination.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listQuotes({ limit, cursor });
        return formatResponse(result, format, 'quotes');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_quote',
    'Get a single quote by ID.',
    { id: z.string(), format: z.enum(['json', 'markdown']).default('json') },
    async ({ id, format }) => {
      try {
        const quote = await client.getQuote(id);
        return formatResponse(quote, format, 'quote');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_quote',
    'Create a new quote in Dynamics 365.',
    {
      name: z.string().describe('Quote name'),
      description: z.string().optional(),
      opportunityId: z.string().optional(),
      customerAccountId: z.string().optional(),
      priceLevelId: z.string().optional(),
      currencyId: z.string().optional(),
      effectiveFrom: z.string().optional(),
      effectiveTo: z.string().optional(),
      discountPercentage: z.number().optional(),
      freightAmount: z.number().optional(),
    },
    async (input) => {
      try {
        const quote = await client.createQuote(input);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Quote created', quote }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_update_quote',
    'Update an existing quote.',
    {
      id: z.string(),
      name: z.string().optional(),
      description: z.string().optional(),
      effectiveFrom: z.string().optional(),
      effectiveTo: z.string().optional(),
      discountPercentage: z.number().optional(),
      freightAmount: z.number().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const quote = await client.updateQuote(id, input);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Quote updated', quote }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_delete_quote',
    'Delete a quote from Dynamics 365.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.deleteQuote(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: `Quote ${id} deleted` }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_list_quote_details',
    'List line items for a quote.',
    { quoteId: z.string() },
    async ({ quoteId }) => {
      try {
        const details = await client.listQuoteDetails(quoteId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, count: details.length, details }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_add_quote_detail',
    'Add a line item to a quote.',
    {
      quoteId: z.string(),
      quantity: z.number(),
      productId: z.string().optional(),
      productDescription: z.string().optional(),
      pricePerUnit: z.number().optional(),
      manualDiscountAmount: z.number().optional(),
      tax: z.number().optional(),
      uomId: z.string().optional(),
      isPriceOverridden: z.boolean().optional(),
    },
    async (input) => {
      try {
        const detail = await client.addQuoteDetail(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Quote detail added', detail }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_activate_quote',
    'Activate a quote (change from draft to active).',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.activateQuote(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Quote activated' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_close_quote',
    'Close a quote with a specific status.',
    {
      id: z.string(),
      status: z.enum(['won', 'lost', 'cancelled']),
    },
    async ({ id, status }) => {
      try {
        await client.closeQuote(id, status);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: `Quote closed as ${status}` }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_convert_quote_to_order',
    'Convert a quote to a sales order.',
    { quoteId: z.string() },
    async ({ quoteId }) => {
      try {
        const result = await client.convertQuoteToOrder(quoteId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Quote converted to order', ...result }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
