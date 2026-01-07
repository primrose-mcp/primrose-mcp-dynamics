/**
 * Product and Price List Tools
 *
 * MCP tools for product catalog management in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

export function registerProductTools(server: McpServer, client: CrmClient): void {
  // Products
  server.tool(
    'dynamics_list_products',
    'List products from the Dynamics 365 product catalog.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listProducts({ limit, cursor });
        return formatResponse(result, format, 'products');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_product',
    'Get a single product by ID.',
    { id: z.string(), format: z.enum(['json', 'markdown']).default('json') },
    async ({ id, format }) => {
      try {
        const product = await client.getProduct(id);
        return formatResponse(product, format, 'product');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_product',
    'Create a new product in the catalog.',
    {
      name: z.string().describe('Product name'),
      productNumber: z.string().describe('Unique product number'),
      description: z.string().optional(),
      productStructure: z.number().int().min(1).max(3).optional().describe('1=Product, 2=Family, 3=Bundle'),
      productTypeCode: z.number().int().optional().describe('1=SalesInventory, 2=MiscCharges, 3=Services, 4=FlatFees'),
      price: z.number().optional(),
      currentCost: z.number().optional(),
      standardCost: z.number().optional(),
      quantityDecimal: z.number().int().optional(),
      defaultUomId: z.string().optional(),
      defaultUomScheduleId: z.string().optional(),
    },
    async (input) => {
      try {
        const product = await client.createProduct(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Product created', product }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_update_product',
    'Update an existing product.',
    {
      id: z.string(),
      name: z.string().optional(),
      productNumber: z.string().optional(),
      description: z.string().optional(),
      price: z.number().optional(),
      currentCost: z.number().optional(),
      standardCost: z.number().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const product = await client.updateProduct(id, input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Product updated', product }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_delete_product',
    'Delete a product from the catalog.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.deleteProduct(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: `Product ${id} deleted` }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_publish_product',
    'Publish a product to make it available for use.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.publishProduct(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Product published' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // Price Lists
  server.tool(
    'dynamics_list_price_lists',
    'List price lists from Dynamics 365.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listPriceLists({ limit, cursor });
        return formatResponse(result, format, 'priceLists');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_price_list',
    'Get a single price list by ID.',
    { id: z.string(), format: z.enum(['json', 'markdown']).default('json') },
    async ({ id, format }) => {
      try {
        const priceList = await client.getPriceList(id);
        return formatResponse(priceList, format, 'priceList');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_price_list',
    'Create a new price list.',
    {
      name: z.string().describe('Price list name'),
      description: z.string().optional(),
      currencyId: z.string().optional(),
    },
    async (input) => {
      try {
        const priceList = await client.createPriceList(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Price list created', priceList }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
