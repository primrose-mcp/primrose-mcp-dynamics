/**
 * Case Tools
 *
 * MCP tools for customer service case management in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

export function registerCaseTools(server: McpServer, client: CrmClient): void {
  server.tool(
    'dynamics_list_cases',
    'List customer service cases from Dynamics 365 with pagination.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listCases({ limit, cursor });
        return formatResponse(result, format, 'cases');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_case',
    'Get a single case by ID.',
    { id: z.string(), format: z.enum(['json', 'markdown']).default('json') },
    async ({ id, format }) => {
      try {
        const caseRecord = await client.getCase(id);
        return formatResponse(caseRecord, format, 'case');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_case',
    'Create a new customer service case.',
    {
      title: z.string().describe('Case title'),
      description: z.string().optional(),
      customerId: z.string().optional().describe('Customer contact or account ID'),
      customerType: z.enum(['contact', 'account']).optional().default('contact'),
      priorityCode: z.number().int().optional().describe('1=High, 2=Normal, 3=Low'),
      caseTypeCode: z.number().int().optional().describe('1=Question, 2=Problem, 3=Request'),
      caseOriginCode: z.number().int().optional().describe('1=Phone, 2=Email, 3=Web, 2483=Social'),
      subjectId: z.string().optional(),
      contractId: z.string().optional(),
      contractDetailId: z.string().optional(),
      productId: z.string().optional(),
      entitlementId: z.string().optional(),
    },
    async (input) => {
      try {
        const caseRecord = await client.createCase(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Case created', case: caseRecord }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_update_case',
    'Update an existing case.',
    {
      id: z.string(),
      title: z.string().optional(),
      description: z.string().optional(),
      priorityCode: z.number().int().optional(),
      caseTypeCode: z.number().int().optional(),
      caseOriginCode: z.number().int().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const caseRecord = await client.updateCase(id, input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Case updated', case: caseRecord }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_delete_case',
    'Delete a case from Dynamics 365.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.deleteCase(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: `Case ${id} deleted` }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_resolve_case',
    'Resolve a case with a resolution description.',
    {
      id: z.string(),
      subject: z.string().describe('Resolution subject'),
      description: z.string().optional().describe('Resolution description'),
      billableTime: z.number().int().optional().describe('Billable time in minutes'),
      resolution: z.string().optional(),
    },
    async ({ id, subject, description, billableTime, resolution }) => {
      try {
        await client.resolveCase(id, { subject, description, billableTime, resolution });
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Case resolved' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_cancel_case',
    'Cancel a case.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.cancelCase(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Case cancelled' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_reactivate_case',
    'Reactivate a resolved or cancelled case.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.reactivateCase(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Case reactivated' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
