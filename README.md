# Microsoft Dynamics 365 MCP Server

[![Primrose MCP](https://img.shields.io/badge/Primrose-MCP-blue)](https://primrose.dev/mcp/dynamics)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A Model Context Protocol (MCP) server for Microsoft Dynamics 365 CRM. Manage contacts, leads, opportunities, cases, and the full suite of CRM entities through a standardized interface.

## Features

- **Activity Management** - Track calls, emails, meetings, and tasks
- **Batch Operations** - Execute multiple operations in a single request
- **Campaign Management** - Create and manage marketing campaigns
- **Case Management** - Handle customer service cases
- **Company/Account Management** - Manage business accounts
- **Competitor Tracking** - Track competitive information
- **Contact Management** - Full CRUD operations for contacts
- **Deal/Opportunity Management** - Track sales opportunities
- **Goal Management** - Set and track sales goals
- **Invoice Processing** - Manage invoices and billing
- **Lead Management** - Capture and qualify leads
- **Metadata Operations** - Access entity metadata
- **Note Management** - Add notes to records
- **Order Management** - Process sales orders
- **Product Catalog** - Manage products and pricing
- **Query Operations** - Execute FetchXML and OData queries
- **Quote Management** - Create and manage quotes
- **Relationship Tracking** - Manage entity relationships
- **User Management** - Access system user information

## Quick Start

The recommended way to use this MCP server is through the [Primrose SDK](https://www.npmjs.com/package/primrose-mcp):

```bash
npm install primrose-mcp
```

```typescript
import { PrimroseClient } from 'primrose-mcp';

const client = new PrimroseClient({
  service: 'dynamics',
  headers: {
    'X-CRM-Access-Token': 'your-access-token',
    'X-Dynamics-Environment-URL': 'https://yourorg.crm.dynamics.com'
  }
});

// List contacts
const contacts = await client.call('dynamics_list_contacts', {
  top: 10
});
```

## Manual Installation

If you prefer to run the MCP server directly:

```bash
# Clone the repository
git clone https://github.com/primrose-ai/primrose-mcp-dynamics.git
cd primrose-mcp-dynamics

# Install dependencies
npm install

# Build the project
npm run build

# Start the server
npm start
```

## Configuration

### Required Headers (one of)

| Header | Description |
|--------|-------------|
| `X-CRM-API-Key` | Your CRM API key |
| `X-CRM-Access-Token` | OAuth access token |

### Optional Headers

| Header | Description |
|--------|-------------|
| `X-CRM-Base-URL` | Override the default CRM API base URL |
| `X-CRM-Client-ID` | OAuth client ID |
| `X-CRM-Client-Secret` | OAuth client secret |
| `X-Dynamics-Tenant-ID` | Microsoft Dynamics tenant ID |
| `X-Dynamics-Environment-URL` | Microsoft Dynamics environment URL |

### Getting Your Access Token

1. Register an application in [Azure Active Directory](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Add Dynamics 365 API permissions
3. Create a client secret
4. Use the OAuth 2.0 client credentials flow to obtain an access token

## Available Tools

### Activity Tools
- `dynamics_list_activities` - List activities
- `dynamics_get_activity` - Get activity details
- `dynamics_create_activity` - Create an activity
- `dynamics_update_activity` - Update an activity
- `dynamics_delete_activity` - Delete an activity

### Batch Tools
- `dynamics_execute_batch` - Execute batch operations

### Campaign Tools
- `dynamics_list_campaigns` - List campaigns
- `dynamics_get_campaign` - Get campaign details
- `dynamics_create_campaign` - Create a campaign
- `dynamics_update_campaign` - Update a campaign

### Case Tools
- `dynamics_list_cases` - List service cases
- `dynamics_get_case` - Get case details
- `dynamics_create_case` - Create a case
- `dynamics_update_case` - Update a case
- `dynamics_resolve_case` - Resolve a case

### Company Tools
- `dynamics_list_accounts` - List accounts
- `dynamics_get_account` - Get account details
- `dynamics_create_account` - Create an account
- `dynamics_update_account` - Update an account

### Contact Tools
- `dynamics_list_contacts` - List contacts
- `dynamics_get_contact` - Get contact details
- `dynamics_create_contact` - Create a contact
- `dynamics_update_contact` - Update a contact
- `dynamics_delete_contact` - Delete a contact

### Deal Tools
- `dynamics_list_opportunities` - List opportunities
- `dynamics_get_opportunity` - Get opportunity details
- `dynamics_create_opportunity` - Create an opportunity
- `dynamics_update_opportunity` - Update an opportunity
- `dynamics_close_opportunity` - Close an opportunity

### Goal Tools
- `dynamics_list_goals` - List goals
- `dynamics_get_goal` - Get goal details
- `dynamics_create_goal` - Create a goal

### Invoice Tools
- `dynamics_list_invoices` - List invoices
- `dynamics_get_invoice` - Get invoice details
- `dynamics_create_invoice` - Create an invoice

### Lead Tools
- `dynamics_list_leads` - List leads
- `dynamics_get_lead` - Get lead details
- `dynamics_create_lead` - Create a lead
- `dynamics_update_lead` - Update a lead
- `dynamics_qualify_lead` - Qualify a lead

### Metadata Tools
- `dynamics_get_entity_metadata` - Get entity metadata
- `dynamics_list_entities` - List entity definitions

### Note Tools
- `dynamics_list_notes` - List notes
- `dynamics_create_note` - Create a note

### Order Tools
- `dynamics_list_orders` - List sales orders
- `dynamics_get_order` - Get order details
- `dynamics_create_order` - Create an order

### Product Tools
- `dynamics_list_products` - List products
- `dynamics_get_product` - Get product details
- `dynamics_create_product` - Create a product

### Query Tools
- `dynamics_execute_fetchxml` - Execute FetchXML query
- `dynamics_odata_query` - Execute OData query

### Quote Tools
- `dynamics_list_quotes` - List quotes
- `dynamics_get_quote` - Get quote details
- `dynamics_create_quote` - Create a quote

### Relationship Tools
- `dynamics_associate` - Associate records
- `dynamics_disassociate` - Disassociate records

### User Tools
- `dynamics_list_users` - List system users
- `dynamics_get_user` - Get user details
- `dynamics_get_current_user` - Get current user

## Development

```bash
# Run in development mode with hot reload
npm run dev

# Run tests
npm test

# Lint code
npm run lint

# Type check
npm run typecheck
```

## Related Resources

- [Primrose SDK Documentation](https://primrose.dev/docs)
- [Dynamics 365 Web API Documentation](https://docs.microsoft.com/en-us/dynamics365/customer-engagement/web-api/about)
- [Model Context Protocol](https://modelcontextprotocol.io)
