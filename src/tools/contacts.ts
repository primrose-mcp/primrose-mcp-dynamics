/**
 * Contact Tools
 *
 * MCP tools for contact management.
 * CUSTOMIZE: Update tool names with your CRM prefix (e.g., hubspot_list_contacts)
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

/**
 * Register all contact-related tools
 *
 * @param server - MCP server instance
 * @param client - CRM client instance
 */
export function registerContactTools(server: McpServer, client: CrmClient): void {
  // ===========================================================================
  // List Contacts
  // ===========================================================================
  server.tool(
    'dynamics_list_contacts',
    `List contacts from the CRM with pagination.

Returns a paginated list of contacts. Use the cursor from the response to fetch the next page.

Args:
  - limit: Number of contacts to return (1-100, default: 20)
  - cursor: Pagination cursor from previous response
  - format: Response format ('json' or 'markdown')

Returns:
  JSON format: { items: Contact[], count, total, hasMore, nextCursor }
  Markdown format: Formatted table of contacts`,
    {
      limit: z.number().int().min(1).max(100).default(20).describe('Number of contacts to return'),
      cursor: z.string().optional().describe('Pagination cursor from previous response'),
      format: z.enum(['json', 'markdown']).default('json').describe('Response format'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listContacts({ limit, cursor });
        return formatResponse(result, format, 'contacts');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Get Contact
  // ===========================================================================
  server.tool(
    'dynamics_get_contact',
    `Get a single contact by ID.

Args:
  - id: The contact ID
  - format: Response format ('json' or 'markdown')

Returns:
  The contact record with all available fields.`,
    {
      id: z.string().describe('Contact ID'),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ id, format }) => {
      try {
        const contact = await client.getContact(id);
        return formatResponse(contact, format, 'contact');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Create Contact
  // ===========================================================================
  server.tool(
    'dynamics_create_contact',
    `Create a new contact in the CRM.

Args:
  - email: Email address (required)
  - firstName: First name
  - lastName: Last name
  - phone: Phone number
  - title: Job title
  - companyId: Associated company ID
  - source: Lead source

Returns:
  The created contact record.`,
    {
      email: z.string().email().describe('Email address (required)'),
      firstName: z.string().optional().describe('First name'),
      lastName: z.string().optional().describe('Last name'),
      phone: z.string().optional().describe('Phone number'),
      title: z.string().optional().describe('Job title'),
      companyId: z.string().optional().describe('Associated company ID'),
      source: z.string().optional().describe('Lead source'),
    },
    async (input) => {
      try {
        const contact = await client.createContact(input);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: 'Contact created', contact }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Update Contact
  // ===========================================================================
  server.tool(
    'dynamics_update_contact',
    `Update an existing contact.

Args:
  - id: Contact ID to update
  - email: New email address
  - firstName: New first name
  - lastName: New last name
  - phone: New phone number
  - title: New job title
  - companyId: New associated company ID

Returns:
  The updated contact record.`,
    {
      id: z.string().describe('Contact ID to update'),
      email: z.string().email().optional(),
      firstName: z.string().optional(),
      lastName: z.string().optional(),
      phone: z.string().optional(),
      title: z.string().optional(),
      companyId: z.string().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const contact = await client.updateContact(id, input);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: 'Contact updated', contact }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Delete Contact
  // ===========================================================================
  server.tool(
    'dynamics_delete_contact',
    `Delete a contact from the CRM.

Args:
  - id: Contact ID to delete

Returns:
  Confirmation of deletion.`,
    {
      id: z.string().describe('Contact ID to delete'),
    },
    async ({ id }) => {
      try {
        await client.deleteContact(id);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: `Contact ${id} deleted` }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Search Contacts
  // ===========================================================================
  server.tool(
    'dynamics_search_contacts',
    `Search for contacts matching criteria.

Args:
  - query: Search query string
  - limit: Maximum results to return
  - format: Response format

Returns:
  Paginated list of matching contacts.`,
    {
      query: z.string().describe('Search query'),
      limit: z.number().int().min(1).max(100).default(20),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ query, limit, format }) => {
      try {
        const result = await client.searchContacts({ query, limit });
        return formatResponse(result, format, 'contacts');
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
