/**
 * Competitor Tools
 *
 * MCP tools for competitor management in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

export function registerCompetitorTools(server: McpServer, client: CrmClient): void {
  server.tool(
    'dynamics_list_competitors',
    'List competitors from Dynamics 365 with pagination.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listCompetitors({ limit, cursor });
        return formatResponse(result, format, 'competitors');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_competitor',
    'Get a single competitor by ID.',
    { id: z.string(), format: z.enum(['json', 'markdown']).default('json') },
    async ({ id, format }) => {
      try {
        const competitor = await client.getCompetitor(id);
        return formatResponse(competitor, format, 'competitor');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_competitor',
    'Create a new competitor in Dynamics 365.',
    {
      name: z.string().describe('Competitor name'),
      websiteUrl: z.string().optional(),
      stockExchange: z.string().optional(),
      tickerSymbol: z.string().optional(),
      strengths: z.string().optional(),
      weaknesses: z.string().optional(),
      threats: z.string().optional(),
      opportunities: z.string().optional(),
      overview: z.string().optional(),
      keyProduct: z.string().optional(),
    },
    async (input) => {
      try {
        const competitor = await client.createCompetitor(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Competitor created', competitor }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_update_competitor',
    'Update an existing competitor.',
    {
      id: z.string(),
      name: z.string().optional(),
      websiteUrl: z.string().optional(),
      stockExchange: z.string().optional(),
      tickerSymbol: z.string().optional(),
      strengths: z.string().optional(),
      weaknesses: z.string().optional(),
      threats: z.string().optional(),
      opportunities: z.string().optional(),
      overview: z.string().optional(),
      keyProduct: z.string().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const competitor = await client.updateCompetitor(id, input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Competitor updated', competitor }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_delete_competitor',
    'Delete a competitor from Dynamics 365.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.deleteCompetitor(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: `Competitor ${id} deleted` }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_associate_competitor_to_opportunity',
    'Associate a competitor with an opportunity.',
    {
      competitorId: z.string(),
      opportunityId: z.string(),
    },
    async ({ competitorId, opportunityId }) => {
      try {
        await client.associateCompetitorToOpportunity(competitorId, opportunityId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Competitor associated with opportunity' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_disassociate_competitor_from_opportunity',
    'Remove a competitor association from an opportunity.',
    {
      competitorId: z.string(),
      opportunityId: z.string(),
    },
    async ({ competitorId, opportunityId }) => {
      try {
        await client.disassociateCompetitorFromOpportunity(competitorId, opportunityId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Competitor disassociated from opportunity' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_list_opportunity_competitors',
    'List competitors associated with an opportunity.',
    { opportunityId: z.string() },
    async ({ opportunityId }) => {
      try {
        const competitors = await client.listOpportunityCompetitors(opportunityId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, count: competitors.length, competitors }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
