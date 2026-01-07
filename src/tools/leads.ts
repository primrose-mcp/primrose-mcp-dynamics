/**
 * Lead Tools
 *
 * MCP tools for lead management in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

/**
 * Register all lead-related tools
 *
 * @param server - MCP server instance
 * @param client - CRM client instance
 */
export function registerLeadTools(server: McpServer, client: CrmClient): void {
  // ===========================================================================
  // List Leads
  // ===========================================================================
  server.tool(
    'dynamics_list_leads',
    `List leads from Dynamics 365 with pagination.

Returns a paginated list of leads. Use the cursor from the response to fetch the next page.

Args:
  - limit: Number of leads to return (1-100, default: 20)
  - cursor: Pagination cursor from previous response
  - format: Response format ('json' or 'markdown')

Returns:
  JSON format: { items: Lead[], count, total, hasMore, nextCursor }
  Markdown format: Formatted table of leads`,
    {
      limit: z.number().int().min(1).max(100).default(20).describe('Number of leads to return'),
      cursor: z.string().optional().describe('Pagination cursor from previous response'),
      format: z.enum(['json', 'markdown']).default('json').describe('Response format'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listLeads({ limit, cursor });
        return formatResponse(result, format, 'leads');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Get Lead
  // ===========================================================================
  server.tool(
    'dynamics_get_lead',
    `Get a single lead by ID.

Args:
  - id: The lead ID (GUID)
  - format: Response format ('json' or 'markdown')

Returns:
  The lead record with all available fields.`,
    {
      id: z.string().describe('Lead ID (GUID)'),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ id, format }) => {
      try {
        const lead = await client.getLead(id);
        return formatResponse(lead, format, 'lead');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Create Lead
  // ===========================================================================
  server.tool(
    'dynamics_create_lead',
    `Create a new lead in Dynamics 365.

Args:
  - lastName: Last name (required)
  - firstName: First name
  - email: Email address
  - phone: Phone number
  - companyName: Company name
  - jobTitle: Job title
  - website: Website URL
  - subject: Lead topic/subject
  - description: Description/notes
  - leadSource: Lead source code (numeric)
  - leadQuality: Lead quality code (1=Hot, 2=Warm, 3=Cold)

Returns:
  The created lead record.`,
    {
      lastName: z.string().describe('Last name (required)'),
      firstName: z.string().optional().describe('First name'),
      email: z.string().email().optional().describe('Email address'),
      phone: z.string().optional().describe('Phone number'),
      companyName: z.string().optional().describe('Company name'),
      jobTitle: z.string().optional().describe('Job title'),
      website: z.string().optional().describe('Website URL'),
      subject: z.string().optional().describe('Lead topic/subject'),
      description: z.string().optional().describe('Description/notes'),
      leadSource: z.number().int().optional().describe('Lead source code'),
      leadQuality: z.number().int().min(1).max(3).optional().describe('Lead quality (1=Hot, 2=Warm, 3=Cold)'),
    },
    async (input) => {
      try {
        const lead = await client.createLead(input);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: 'Lead created', lead }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Update Lead
  // ===========================================================================
  server.tool(
    'dynamics_update_lead',
    `Update an existing lead.

Args:
  - id: Lead ID to update (required)
  - lastName: New last name
  - firstName: New first name
  - email: New email address
  - phone: New phone number
  - companyName: New company name
  - jobTitle: New job title
  - website: New website URL
  - subject: New lead topic/subject
  - description: New description/notes
  - leadSource: New lead source code
  - leadQuality: New lead quality (1=Hot, 2=Warm, 3=Cold)

Returns:
  The updated lead record.`,
    {
      id: z.string().describe('Lead ID to update'),
      lastName: z.string().optional(),
      firstName: z.string().optional(),
      email: z.string().email().optional(),
      phone: z.string().optional(),
      companyName: z.string().optional(),
      jobTitle: z.string().optional(),
      website: z.string().optional(),
      subject: z.string().optional(),
      description: z.string().optional(),
      leadSource: z.number().int().optional(),
      leadQuality: z.number().int().min(1).max(3).optional(),
    },
    async ({ id, ...input }) => {
      try {
        const lead = await client.updateLead(id, input);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: 'Lead updated', lead }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Delete Lead
  // ===========================================================================
  server.tool(
    'dynamics_delete_lead',
    `Delete a lead from Dynamics 365.

Args:
  - id: Lead ID to delete

Returns:
  Confirmation of deletion.`,
    {
      id: z.string().describe('Lead ID to delete'),
    },
    async ({ id }) => {
      try {
        await client.deleteLead(id);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: `Lead ${id} deleted` }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Search Leads
  // ===========================================================================
  server.tool(
    'dynamics_search_leads',
    `Search for leads matching criteria.

Searches across full name, email address, and company name.

Args:
  - query: Search query string
  - limit: Maximum results to return
  - format: Response format

Returns:
  Paginated list of matching leads.`,
    {
      query: z.string().describe('Search query'),
      limit: z.number().int().min(1).max(100).default(20),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ query, limit, format }) => {
      try {
        const result = await client.searchLeads({ query, limit });
        return formatResponse(result, format, 'leads');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Qualify Lead
  // ===========================================================================
  server.tool(
    'dynamics_qualify_lead',
    `Qualify a lead, optionally creating an Account, Contact, and/or Opportunity.

This action converts a lead to qualified status and can automatically create
related records based on the lead data.

Args:
  - id: Lead ID to qualify (required)
  - createAccount: Whether to create an Account (default: true)
  - createContact: Whether to create a Contact (default: true)
  - createOpportunity: Whether to create an Opportunity (default: true)
  - opportunityCurrencyId: Currency ID for the opportunity
  - opportunityCustomerId: Customer ID for the opportunity

Returns:
  Object containing IDs of created records (accountId, contactId, opportunityId).`,
    {
      id: z.string().describe('Lead ID to qualify'),
      createAccount: z.boolean().default(true).describe('Create an Account from the lead'),
      createContact: z.boolean().default(true).describe('Create a Contact from the lead'),
      createOpportunity: z.boolean().default(true).describe('Create an Opportunity from the lead'),
      opportunityCurrencyId: z.string().optional().describe('Currency ID for the opportunity'),
      opportunityCustomerId: z.string().optional().describe('Customer ID for the opportunity'),
    },
    async ({ id, createAccount, createContact, createOpportunity, opportunityCurrencyId, opportunityCustomerId }) => {
      try {
        const result = await client.qualifyLead(id, {
          createAccount,
          createContact,
          createOpportunity,
          opportunityCurrencyId,
          opportunityCustomerId,
        });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  message: 'Lead qualified successfully',
                  ...result,
                  success: true,
                },
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
}
