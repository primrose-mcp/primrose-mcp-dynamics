/**
 * Company Tools
 *
 * MCP tools for company/account management.
 * CUSTOMIZE: Update tool names with your CRM prefix (e.g., hubspot_list_companies)
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

/**
 * Register all company-related tools
 */
export function registerCompanyTools(server: McpServer, client: CrmClient): void {
  // ===========================================================================
  // List Companies
  // ===========================================================================
  server.tool(
    'dynamics_list_companies',
    `List companies from the CRM with pagination.

Args:
  - limit: Number of companies to return (1-100, default: 20)
  - cursor: Pagination cursor from previous response
  - format: Response format ('json' or 'markdown')

Returns:
  Paginated list of companies.`,
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listCompanies({ limit, cursor });
        return formatResponse(result, format, 'companies');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Get Company
  // ===========================================================================
  server.tool(
    'dynamics_get_company',
    `Get a single company by ID.

Args:
  - id: The company ID
  - format: Response format

Returns:
  The company record with all available fields.`,
    {
      id: z.string().describe('Company ID'),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ id, format }) => {
      try {
        const company = await client.getCompany(id);
        return formatResponse(company, format, 'company');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Create Company
  // ===========================================================================
  server.tool(
    'dynamics_create_company',
    `Create a new company in the CRM.

Args:
  - name: Company name (required)
  - domain: Company domain/website
  - industry: Industry classification
  - description: Company description
  - numberOfEmployees: Number of employees
  - type: Company type (prospect, customer, partner, etc.)
  - phone: Phone number

Returns:
  The created company record.`,
    {
      name: z.string().describe('Company name (required)'),
      domain: z.string().optional().describe('Company domain/website'),
      industry: z.string().optional().describe('Industry classification'),
      description: z.string().optional().describe('Company description'),
      numberOfEmployees: z.number().int().optional().describe('Number of employees'),
      type: z.string().optional().describe('Company type'),
      phone: z.string().optional().describe('Phone number'),
    },
    async (input) => {
      try {
        const company = await client.createCompany(input);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: 'Company created', company }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Update Company
  // ===========================================================================
  server.tool(
    'dynamics_update_company',
    `Update an existing company.

Args:
  - id: Company ID to update
  - name: New company name
  - domain: New domain
  - industry: New industry
  - description: New description
  - numberOfEmployees: New employee count
  - type: New company type
  - phone: New phone number

Returns:
  The updated company record.`,
    {
      id: z.string().describe('Company ID to update'),
      name: z.string().optional(),
      domain: z.string().optional(),
      industry: z.string().optional(),
      description: z.string().optional(),
      numberOfEmployees: z.number().int().optional(),
      type: z.string().optional(),
      phone: z.string().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const company = await client.updateCompany(id, input);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: 'Company updated', company }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
