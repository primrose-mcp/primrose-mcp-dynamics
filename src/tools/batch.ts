/**
 * Batch Operations Tools
 *
 * MCP tools for batch/bulk operations in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError } from '../utils/formatters.js';

export function registerBatchTools(server: McpServer, client: CrmClient): void {
  server.tool(
    'dynamics_batch_create',
    'Create multiple records in a single batch request. Supports up to 1000 records per batch.',
    {
      entityType: z.string().describe('Entity logical name (e.g., contact, account, lead)'),
      records: z.array(z.record(z.string(), z.unknown())).describe('Array of records to create'),
      continueOnError: z.boolean().optional().default(false).describe('Continue processing on error'),
    },
    async ({ entityType, records, continueOnError }) => {
      try {
        const results = await client.batchCreate(entityType, records, continueOnError);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              message: `Batch created ${results.succeeded} records`,
              succeeded: results.succeeded,
              failed: results.failed,
              errors: results.errors,
              createdIds: results.createdIds,
            }, null, 2),
          }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_batch_update',
    'Update multiple records in a single batch request.',
    {
      entityType: z.string().describe('Entity logical name'),
      updates: z.array(z.object({
        id: z.string().describe('Record ID'),
        data: z.record(z.string(), z.unknown()).describe('Fields to update'),
      })).describe('Array of updates'),
      continueOnError: z.boolean().optional().default(false),
    },
    async ({ entityType, updates, continueOnError }) => {
      try {
        const results = await client.batchUpdate(entityType, updates, continueOnError);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              message: `Batch updated ${results.succeeded} records`,
              succeeded: results.succeeded,
              failed: results.failed,
              errors: results.errors,
            }, null, 2),
          }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_batch_delete',
    'Delete multiple records in a single batch request.',
    {
      entityType: z.string().describe('Entity logical name'),
      ids: z.array(z.string()).describe('Array of record IDs to delete'),
      continueOnError: z.boolean().optional().default(false),
    },
    async ({ entityType, ids, continueOnError }) => {
      try {
        const results = await client.batchDelete(entityType, ids, continueOnError);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              message: `Batch deleted ${results.succeeded} records`,
              succeeded: results.succeeded,
              failed: results.failed,
              errors: results.errors,
            }, null, 2),
          }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_batch_upsert',
    'Create or update multiple records based on alternate key or ID.',
    {
      entityType: z.string().describe('Entity logical name'),
      records: z.array(z.object({
        alternateKey: z.record(z.string(), z.string()).optional().describe('Alternate key fields'),
        id: z.string().optional().describe('Record ID (if updating by ID)'),
        data: z.record(z.string(), z.unknown()).describe('Record data'),
      })).describe('Array of records to upsert'),
      continueOnError: z.boolean().optional().default(false),
    },
    async ({ entityType, records, continueOnError }) => {
      try {
        const results = await client.batchUpsert(entityType, records, continueOnError);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              message: `Batch upserted ${results.succeeded} records`,
              succeeded: results.succeeded,
              failed: results.failed,
              created: results.created,
              updated: results.updated,
              errors: results.errors,
            }, null, 2),
          }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
