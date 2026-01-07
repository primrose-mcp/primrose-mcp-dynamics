/**
 * Deal Tools
 *
 * MCP tools for deal/opportunity management.
 * CUSTOMIZE: Update tool names with your CRM prefix (e.g., hubspot_list_deals)
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

/**
 * Register all deal-related tools
 */
export function registerDealTools(server: McpServer, client: CrmClient): void {
  // ===========================================================================
  // List Deals
  // ===========================================================================
  server.tool(
    'dynamics_list_deals',
    `List deals/opportunities from the CRM with pagination.

Args:
  - limit: Number of deals to return (1-100, default: 20)
  - cursor: Pagination cursor from previous response
  - format: Response format ('json' or 'markdown')

Returns:
  Paginated list of deals.`,
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listDeals({ limit, cursor });
        return formatResponse(result, format, 'deals');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Get Deal
  // ===========================================================================
  server.tool(
    'dynamics_get_deal',
    `Get a single deal by ID.

Args:
  - id: The deal ID
  - format: Response format

Returns:
  The deal record with all available fields.`,
    {
      id: z.string().describe('Deal ID'),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ id, format }) => {
      try {
        const deal = await client.getDeal(id);
        return formatResponse(deal, format, 'deal');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Create Deal
  // ===========================================================================
  server.tool(
    'dynamics_create_deal',
    `Create a new deal in the CRM.

Args:
  - name: Deal name (required)
  - amount: Deal value
  - currency: Currency code (e.g., 'USD', 'AUD')
  - stageId: Pipeline stage ID
  - pipelineId: Pipeline ID
  - closeDate: Expected close date (ISO 8601)
  - companyId: Associated company ID
  - contactIds: Associated contact IDs

Returns:
  The created deal record.`,
    {
      name: z.string().describe('Deal name (required)'),
      amount: z.number().optional().describe('Deal value'),
      currency: z.string().optional().describe('Currency code'),
      stageId: z.string().optional().describe('Pipeline stage ID'),
      pipelineId: z.string().optional().describe('Pipeline ID'),
      closeDate: z.string().optional().describe('Expected close date (ISO 8601)'),
      companyId: z.string().optional().describe('Associated company ID'),
      contactIds: z.array(z.string()).optional().describe('Associated contact IDs'),
    },
    async (input) => {
      try {
        const deal = await client.createDeal(input);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: 'Deal created', deal }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Update Deal
  // ===========================================================================
  server.tool(
    'dynamics_update_deal',
    `Update an existing deal.

Args:
  - id: Deal ID to update
  - name: New deal name
  - amount: New deal value
  - currency: New currency code
  - stageId: New pipeline stage ID
  - closeDate: New expected close date
  - status: Deal status ('open', 'won', 'lost')
  - closeReason: Reason for won/lost

Returns:
  The updated deal record.`,
    {
      id: z.string().describe('Deal ID to update'),
      name: z.string().optional(),
      amount: z.number().optional(),
      currency: z.string().optional(),
      stageId: z.string().optional(),
      closeDate: z.string().optional(),
      status: z.enum(['open', 'won', 'lost']).optional(),
      closeReason: z.string().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const deal = await client.updateDeal(id, input);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: 'Deal updated', deal }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Move Deal Stage
  // ===========================================================================
  server.tool(
    'dynamics_move_deal_stage',
    `Move a deal to a different pipeline stage.

Args:
  - id: Deal ID
  - stageId: Target stage ID

Returns:
  The updated deal record.`,
    {
      id: z.string().describe('Deal ID'),
      stageId: z.string().describe('Target stage ID'),
    },
    async ({ id, stageId }) => {
      try {
        const deal = await client.moveDealStage(id, stageId);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                { success: true, message: `Deal moved to stage ${stageId}`, deal },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // List Pipelines
  // ===========================================================================
  server.tool(
    'dynamics_list_pipelines',
    `List available deal pipelines and their stages.

Returns:
  Array of pipelines with their stages.`,
    {
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ format }) => {
      try {
        const pipelines = await client.listPipelines();
        return formatResponse(pipelines, format, 'pipelines');
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
