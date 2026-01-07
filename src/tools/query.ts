/**
 * OData Query Tools
 *
 * MCP tools for advanced OData queries in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError } from '../utils/formatters.js';

export function registerQueryTools(server: McpServer, client: CrmClient): void {
  server.tool(
    'dynamics_query',
    `Execute an advanced OData query against any Dynamics 365 entity.

Supports:
- $select: Choose specific fields
- $filter: Filter conditions (e.g., "statecode eq 0", "contains(name,'Acme')")
- $orderby: Sort results
- $expand: Include related records
- $top/$skip: Pagination

Example filters:
- statecode eq 0 (active records)
- contains(name,'Microsoft')
- createdon ge 2024-01-01
- revenue gt 1000000
- _parentcustomerid_value eq {guid}`,
    {
      entityType: z.string().describe('Entity logical name (e.g., contact, account, opportunity)'),
      select: z.string().optional().describe('Comma-separated fields to return'),
      filter: z.string().optional().describe('OData filter expression'),
      orderby: z.string().optional().describe('Sort expression (e.g., "createdon desc")'),
      expand: z.string().optional().describe('Related entities to include'),
      top: z.number().int().min(1).max(5000).optional().default(50).describe('Max records to return'),
      skip: z.number().int().min(0).optional().describe('Records to skip'),
      count: z.boolean().optional().default(false).describe('Include total count'),
    },
    async ({ entityType, select, filter, orderby, expand, top, skip, count }) => {
      try {
        const result = await client.executeQuery({
          entityType,
          select,
          filter,
          orderby,
          expand,
          top,
          skip,
          count,
        });
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              count: result.records.length,
              total: result.totalCount,
              hasMore: result.hasMore,
              records: result.records,
            }, null, 2),
          }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_aggregate',
    'Execute an aggregate query (count, sum, avg, min, max) on an entity.',
    {
      entityType: z.string().describe('Entity logical name'),
      aggregateFunction: z.enum(['count', 'sum', 'avg', 'min', 'max']).describe('Aggregate function'),
      field: z.string().optional().describe('Field to aggregate (not needed for count)'),
      filter: z.string().optional().describe('OData filter expression'),
      groupBy: z.string().optional().describe('Field to group by'),
    },
    async ({ entityType, aggregateFunction, field, filter, groupBy }) => {
      try {
        const result = await client.executeAggregate({
          entityType,
          aggregateFunction,
          field,
          filter,
          groupBy,
        });
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, result }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_count',
    'Get the count of records matching a filter.',
    {
      entityType: z.string().describe('Entity logical name'),
      filter: z.string().optional().describe('OData filter expression'),
    },
    async ({ entityType, filter }) => {
      try {
        const count = await client.getRecordCount(entityType, filter);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, count }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_saved_query',
    'Execute a saved (system) view query.',
    {
      savedQueryId: z.string().describe('Saved query ID'),
      entityType: z.string().describe('Entity logical name'),
      top: z.number().int().min(1).max(5000).optional().default(50),
    },
    async ({ savedQueryId, entityType, top }) => {
      try {
        const result = await client.executeSavedQuery(savedQueryId, entityType, top);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              count: result.length,
              records: result,
            }, null, 2),
          }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_user_query',
    'Execute a user (personal) view query.',
    {
      userQueryId: z.string().describe('User query ID'),
      entityType: z.string().describe('Entity logical name'),
      top: z.number().int().min(1).max(5000).optional().default(50),
    },
    async ({ userQueryId, entityType, top }) => {
      try {
        const result = await client.executeUserQuery(userQueryId, entityType, top);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              count: result.length,
              records: result,
            }, null, 2),
          }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
