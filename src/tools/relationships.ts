/**
 * Relationship Tools
 *
 * MCP tools for managing entity relationships in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError } from '../utils/formatters.js';

export function registerRelationshipTools(server: McpServer, client: CrmClient): void {
  server.tool(
    'dynamics_associate_records',
    'Associate two records via a relationship.',
    {
      sourceEntityType: z.string().describe('Source entity type (e.g., contact)'),
      sourceId: z.string().describe('Source record ID'),
      targetEntityType: z.string().describe('Target entity type (e.g., account)'),
      targetId: z.string().describe('Target record ID'),
      relationshipName: z.string().describe('Relationship schema name'),
    },
    async ({ sourceEntityType, sourceId, targetEntityType, targetId, relationshipName }) => {
      try {
        await client.associateRecords(sourceEntityType, sourceId, targetEntityType, targetId, relationshipName);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Records associated successfully' }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_disassociate_records',
    'Remove an association between two records.',
    {
      sourceEntityType: z.string().describe('Source entity type'),
      sourceId: z.string().describe('Source record ID'),
      targetEntityType: z.string().describe('Target entity type'),
      targetId: z.string().describe('Target record ID'),
      relationshipName: z.string().describe('Relationship schema name'),
    },
    async ({ sourceEntityType, sourceId, targetEntityType, targetId, relationshipName }) => {
      try {
        await client.disassociateRecords(sourceEntityType, sourceId, targetEntityType, targetId, relationshipName);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Records disassociated successfully' }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_list_related_records',
    'List records related to a source record via a navigation property.',
    {
      entityType: z.string().describe('Source entity type'),
      id: z.string().describe('Source record ID'),
      navigationProperty: z.string().describe('Navigation property name'),
      select: z.string().optional().describe('Comma-separated list of fields to select'),
      limit: z.number().int().min(1).max(100).optional().default(20),
    },
    async ({ entityType, id, navigationProperty, select, limit }) => {
      try {
        const records = await client.listRelatedRecords(entityType, id, navigationProperty, { select, limit });
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              count: records.length,
              records,
            }, null, 2),
          }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
