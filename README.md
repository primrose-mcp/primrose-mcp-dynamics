# Primrose MCP Dynamics 365

A Model Context Protocol (MCP) server for Microsoft Dynamics 365, deployed on Cloudflare Workers.

## Features

- **Contacts** - List, get, create, update, delete, and search contacts
- **Companies (Accounts)** - Full CRUD operations for accounts
- **Deals (Opportunities)** - Opportunity management with sales stages
- **Leads** - Lead qualification and management
- **Cases** - Customer service case management
- **Campaigns** - Marketing campaign management
- **Products** - Product catalog management
- **Quotes** - Quote creation and management
- **Orders** - Order processing
- **Invoices** - Invoice management
- **Activities** - Phone calls, emails, appointments, and tasks
- **Notes** - Annotation management
- **Competitors** - Competitor tracking
- **Goals** - Sales goal management
- **Relationships** - Connection and relationship tracking
- **Users** - System user management
- **Metadata** - Entity and attribute metadata
- **Query** - FetchXML and OData queries
- **Batch** - Batch and transaction operations

## Authentication

This server uses a multi-tenant architecture. Pass your Dynamics 365 credentials via request headers:

| Header | Description |
|--------|-------------|
| `X-CRM-Access-Token` | Azure AD OAuth access token (required) |
| `X-CRM-Base-URL` | Your Dynamics 365 instance URL (required) |

## Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/mcp` | POST | Streamable HTTP MCP endpoint |
| `/health` | GET | Health check |

## Installation

```bash
npm install
```

## Development

```bash
# Local development
npm run dev

# Type checking
npm run typecheck

# Linting
npm run lint
npm run lint:fix

# Format code
npm run format
```

## Deployment

```bash
npm run deploy
```

## Testing

Use the [MCP Inspector](https://github.com/modelcontextprotocol/inspector) to test the server:

```bash
npx @modelcontextprotocol/inspector
```

## API Reference

- [Dynamics 365 Web API Documentation](https://learn.microsoft.com/en-us/dynamics365/customerengagement/on-premises/developer/webapi/web-api-versions)
- API Version: v9.2

## License

MIT
