/**
 * CRM MCP Server - Main Entry Point
 *
 * This file sets up the MCP server using Cloudflare's Agents SDK.
 * It supports both stateless (McpServer) and stateful (McpAgent) modes.
 *
 * MULTI-TENANT ARCHITECTURE:
 * Tenant credentials (API keys, etc.) are parsed from request headers,
 * allowing a single server deployment to serve multiple customers.
 *
 * Required Headers:
 * - X-CRM-API-Key: API key for CRM authentication
 *
 * Optional Headers:
 * - X-CRM-Base-URL: Override the default CRM API base URL
 * - X-CRM-Access-Token: OAuth access token (alternative to API key)
 *
 * CUSTOMIZE: Update the server name, version, and register your tools.
 */

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { McpAgent } from 'agents/mcp';
import { z } from 'zod';
import { createCrmClient } from './client.js';
import { registerActivityTools } from './tools/activities.js';
import { registerBatchTools } from './tools/batch.js';
import { registerCampaignTools } from './tools/campaigns.js';
import { registerCaseTools } from './tools/cases.js';
import { registerCompanyTools } from './tools/companies.js';
import { registerCompetitorTools } from './tools/competitors.js';
import { registerContactTools } from './tools/contacts.js';
import { registerDealTools } from './tools/deals.js';
import { registerGoalTools } from './tools/goals.js';
import { registerInvoiceTools } from './tools/invoices.js';
import { registerLeadTools } from './tools/leads.js';
import { registerMetadataTools } from './tools/metadata.js';
import { registerNoteTools } from './tools/notes.js';
import { registerOrderTools } from './tools/orders.js';
import { registerProductTools } from './tools/products.js';
import { registerQueryTools } from './tools/query.js';
import { registerQuoteTools } from './tools/quotes.js';
import { registerRelationshipTools } from './tools/relationships.js';
import { registerUserTools } from './tools/users.js';
import {
  type Env,
  type TenantCredentials,
  parseTenantCredentials,
  validateCredentials,
} from './types/env.js';

// =============================================================================
// MCP Server Configuration
// =============================================================================

/**
 * CUSTOMIZE: Update these values for your CRM
 */
const SERVER_NAME = 'lineer-mcp-dynamics';
const SERVER_VERSION = '1.0.0';

// =============================================================================
// MCP Agent (Stateful - uses Durable Objects)
// =============================================================================

/**
 * McpAgent provides stateful MCP sessions backed by Durable Objects.
 *
 * NOTE: For multi-tenant deployments, use the stateless mode (Option 2) instead.
 * The stateful McpAgent is better suited for single-tenant deployments where
 * credentials can be stored as wrangler secrets.
 *
 * Use this when you need:
 * - Session state persistence
 * - Per-user rate limiting
 * - Cached API responses
 *
 * @deprecated For multi-tenant support, use stateless mode with per-request credentials
 */
export class CrmMcpAgent extends McpAgent<Env> {
  server = new McpServer({
    name: SERVER_NAME,
    version: SERVER_VERSION,
  });

  async init() {
    // NOTE: Stateful mode requires credentials to be configured differently.
    // For multi-tenant, use the stateless endpoint at /mcp instead.
    throw new Error(
      'Stateful mode (McpAgent) is not supported for multi-tenant deployments. ' +
        'Use the stateless /mcp endpoint with X-CRM-API-Key header instead.'
    );
  }
}

// =============================================================================
// Stateless MCP Server (Recommended - no Durable Objects needed)
// =============================================================================

/**
 * Creates a stateless MCP server instance with tenant-specific credentials.
 *
 * MULTI-TENANT: Each request provides credentials via headers, allowing
 * a single server deployment to serve multiple tenants.
 *
 * @param credentials - Tenant credentials parsed from request headers
 */
function createStatelessServer(credentials: TenantCredentials): McpServer {
  const server = new McpServer({
    name: SERVER_NAME,
    version: SERVER_VERSION,
  });

  // Create client with tenant-specific credentials
  const client = createCrmClient(credentials);

  // Core CRM entities
  registerContactTools(server, client);
  registerCompanyTools(server, client);
  registerDealTools(server, client);
  registerActivityTools(server, client);
  registerLeadTools(server, client);

  // Sales process entities
  registerQuoteTools(server, client);
  registerOrderTools(server, client);
  registerInvoiceTools(server, client);
  registerProductTools(server, client);

  // Marketing and service entities
  registerCompetitorTools(server, client);
  registerCampaignTools(server, client);
  registerCaseTools(server, client);

  // Management entities
  registerGoalTools(server, client);
  registerNoteTools(server, client);
  registerUserTools(server, client);

  // Advanced operations
  registerBatchTools(server, client);
  registerQueryTools(server, client);
  registerMetadataTools(server, client);
  registerRelationshipTools(server, client);

  server.tool('dynamics_test_connection', 'Test the connection to Dynamics 365 Web API', {}, async () => {
    try {
      const result = await client.testConnection();
      return {
        content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Connection failed: ${error instanceof Error ? error.message : 'Unknown error'}`,
          },
        ],
        isError: true,
      };
    }
  });

  // ===========================================================================
  // Dynamics-specific: FetchXML Query
  // ===========================================================================
  server.tool(
    'dynamics_fetchxml_query',
    `Execute a FetchXML query against Dynamics 365.

FetchXML is a proprietary query language used in Dynamics 365 for complex queries.
This tool allows you to run advanced queries that cannot be expressed with standard OData.

Args:
  - entity: Entity logical name (e.g., 'contact', 'account', 'opportunity')
  - fetchxml: FetchXML query string

Example FetchXML:
<fetch top="10">
  <entity name="contact">
    <attribute name="fullname"/>
    <attribute name="emailaddress1"/>
    <filter>
      <condition attribute="statecode" operator="eq" value="0"/>
    </filter>
  </entity>
</fetch>

Returns:
  Array of matching records.`,
    {
      entity: z.string().describe('Entity logical name (e.g., contact, account, opportunity)'),
      fetchxml: z.string().describe('FetchXML query string'),
    },
    async ({ entity, fetchxml }) => {
      try {
        const results = await client.executeFetchXml(entity, fetchxml);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, count: results.length, results }, null, 2),
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: `FetchXML query failed: ${error instanceof Error ? error.message : 'Unknown error'}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  return server;
}

// =============================================================================
// Worker Export
// =============================================================================

export default {
  /**
   * Main fetch handler for the Worker
   */
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    const url = new URL(request.url);

    // Health check endpoint
    if (url.pathname === '/health') {
      return new Response(JSON.stringify({ status: 'ok', server: SERVER_NAME }), {
        headers: { 'Content-Type': 'application/json' },
      });
    }

    // ==========================================================================
    // Option 1: Stateful MCP with McpAgent (requires Durable Objects)
    // ==========================================================================
    // Uncomment to use McpAgent for stateful sessions:
    //
    // if (url.pathname === '/sse' || url.pathname === '/mcp') {
    //   return CrmMcpAgent.serveSSE('/sse').fetch(request, env, ctx);
    // }

    // ==========================================================================
    // Option 2: Stateless MCP with Streamable HTTP (Recommended for multi-tenant)
    // ==========================================================================
    if (url.pathname === '/mcp' && request.method === 'POST') {
      // Parse tenant credentials from request headers
      const credentials = parseTenantCredentials(request);

      // Validate credentials are present
      try {
        validateCredentials(credentials);
      } catch (error) {
        return new Response(
          JSON.stringify({
            error: 'Unauthorized',
            message: error instanceof Error ? error.message : 'Invalid credentials',
            required_headers: {
              'X-Dynamics-Environment-URL': 'Your Dynamics environment URL (required)',
              'X-CRM-Access-Token':
                'OAuth access token (if using pre-obtained token)',
              'X-Dynamics-Tenant-ID': 'Azure AD tenant ID (if using client credentials)',
              'X-CRM-Client-ID': 'Azure AD app client ID (if using client credentials)',
              'X-CRM-Client-Secret':
                'Azure AD app client secret (if using client credentials)',
            },
          }),
          {
            status: 401,
            headers: { 'Content-Type': 'application/json' },
          }
        );
      }

      // Create server with tenant-specific credentials
      const server = createStatelessServer(credentials);

      // Import and use createMcpHandler for streamable HTTP
      // This is the recommended approach for stateless MCP servers
      const { createMcpHandler } = await import('agents/mcp');
      const handler = createMcpHandler(server);
      return handler(request, env, ctx);
    }

    // SSE endpoint for legacy clients
    if (url.pathname === '/sse') {
      // For SSE, we need to use McpAgent with serveSSE
      // Enable Durable Objects in wrangler.jsonc to use this
      return new Response('SSE endpoint requires Durable Objects. Enable in wrangler.jsonc.', {
        status: 501,
      });
    }

    // Default response
    return new Response(
      JSON.stringify({
        name: SERVER_NAME,
        version: SERVER_VERSION,
        description: 'Microsoft Dynamics 365 CRM MCP Server (Multi-tenant)',
        endpoints: {
          mcp: '/mcp (POST) - Streamable HTTP MCP endpoint',
          health: '/health - Health check',
        },
        authentication: {
          description: 'Pass tenant credentials via request headers',
          option1_access_token: {
            'X-Dynamics-Environment-URL':
              'Your Dynamics environment URL (e.g., https://yourorg.crm.dynamics.com)',
            'X-CRM-Access-Token': 'Pre-obtained OAuth access token',
          },
          option2_client_credentials: {
            'X-Dynamics-Environment-URL':
              'Your Dynamics environment URL (e.g., https://yourorg.crm.dynamics.com)',
            'X-Dynamics-Tenant-ID': 'Azure AD tenant ID',
            'X-CRM-Client-ID': 'Azure AD app client ID',
            'X-CRM-Client-Secret': 'Azure AD app client secret',
          },
        },
      }),
      {
        headers: { 'Content-Type': 'application/json' },
      }
    );
  },
};
