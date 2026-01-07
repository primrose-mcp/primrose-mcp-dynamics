# CRM MCP Server Template

A production-ready template for building MCP (Model Context Protocol) servers that integrate with CRM systems, deployed on Cloudflare Workers.

## Overview

This template provides the foundation for creating **individual MCP server repositories** for each CRM integration:

- `hubspot-mcp-server`
- `salesforce-mcp-server`
- `dynamics-mcp-server`
- `pipedrive-mcp-server`
- `monday-crm-mcp-server`

Each CRM gets its own repository, codebase, and deployment - this template provides the common structure and patterns.

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                     MCP Client (Claude, etc.)                    │
└─────────────────────────┬───────────────────────────────────────┘
                          │ Streamable HTTP / SSE
                          ▼
┌─────────────────────────────────────────────────────────────────┐
│                   Cloudflare Worker                              │
│  ┌─────────────────────────────────────────────────────────────┐│
│  │                 McpAgent (from agents SDK)                   ││
│  │  ┌─────────────┐ ┌─────────────┐ ┌─────────────────────────┐││
│  │  │   Tools     │ │  Resources  │ │  OAuth (optional)       │││
│  │  └──────┬──────┘ └──────┬──────┘ └──────────┬──────────────┘││
│  └─────────┼───────────────┼───────────────────┼───────────────┘│
│            └───────────────┴───────────────────┘                 │
│                            │                                     │
│  ┌─────────────────────────┴───────────────────────────────────┐│
│  │                   CRM API Client                             ││
│  │  - Authentication    - Rate Limiting    - Error Handling     ││
│  └─────────────────────────────────────────────────────────────┘│
└─────────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────────┐
│                      CRM API (e.g., HubSpot)                     │
└─────────────────────────────────────────────────────────────────┘
```

## Quick Start

### 1. Copy this template

```bash
cp -r crm-mcp-template hubspot-mcp-server
cd hubspot-mcp-server
```

### 2. Update configuration

```bash
# Update package.json name
sed -i 's/crm-mcp-server/hubspot-mcp-server/g' package.json

# Update wrangler.jsonc
sed -i 's/crm-mcp-server/hubspot-mcp-server/g' wrangler.jsonc
```

### 3. Install dependencies

```bash
# Using bun (recommended)
bun install

# Or using pnpm
pnpm install

# Or using npm
npm install
```

### 4. Implement CRM-specific logic

- Update `src/client.ts` with CRM API endpoints
- Implement tools in `src/tools/`
- Add CRM-specific types in `src/types/`

### 5. Deploy

```bash
# Using bun
bun run deploy

# Or using pnpm/npm
pnpm run deploy
npm run deploy
```

## Project Structure

```
{crm}-mcp-server/
├── src/
│   ├── index.ts              # Worker entry point & MCP server setup
│   ├── client.ts             # CRM API client (customize per CRM)
│   ├── tools/
│   │   ├── index.ts          # Tool registration
│   │   ├── contacts.ts       # Contact CRUD tools
│   │   ├── companies.ts      # Company CRUD tools
│   │   ├── deals.ts          # Deal/Opportunity tools
│   │   └── activities.ts     # Activity logging tools
│   ├── types/
│   │   ├── env.ts            # Environment bindings
│   │   └── entities.ts       # CRM entity types
│   └── utils/
│       ├── errors.ts         # Error handling
│       ├── pagination.ts     # Pagination helpers
│       └── formatters.ts     # Response formatting
├── package.json
├── tsconfig.json
├── wrangler.jsonc
├── IMPLEMENTATION_GUIDE.md   # Ralph Wiggum instructions
└── README.md
```

## Using with Ralph Wiggum

This template is designed for iterative development with Claude Code's Ralph Wiggum plugin.

### Example: Building HubSpot MCP Server

```bash
/ralph-loop "Build a HubSpot MCP Server from the crm-mcp-template.

Phase 1 - Setup:
- Copy template to hubspot-mcp-server
- Update package.json and wrangler.jsonc names
- Install dependencies

Phase 2 - API Client:
- Implement HubSpot API client in src/client.ts
- Add OAuth and API key authentication
- Add rate limiting (100 req/10s)

Phase 3 - Tools:
- Implement contact tools (list, get, create, update, delete)
- Implement company tools
- Implement deal tools with pipeline stages
- Implement activity logging (calls, emails)

Phase 4 - Testing:
- Run npm run build (must pass)
- Test with MCP Inspector
- Deploy to Cloudflare Workers

Output <promise>COMPLETE</promise> when deployed and tested." --max-iterations 40
```

## Standard Tool Set

Each CRM MCP server should implement these tools:

### Contacts
| Tool | Description |
|------|-------------|
| `{crm}_list_contacts` | List contacts with pagination |
| `{crm}_get_contact` | Get contact by ID |
| `{crm}_create_contact` | Create new contact |
| `{crm}_update_contact` | Update existing contact |
| `{crm}_delete_contact` | Delete contact |
| `{crm}_search_contacts` | Search contacts |

### Companies
| Tool | Description |
|------|-------------|
| `{crm}_list_companies` | List companies with pagination |
| `{crm}_get_company` | Get company by ID |
| `{crm}_create_company` | Create new company |
| `{crm}_update_company` | Update existing company |

### Deals
| Tool | Description |
|------|-------------|
| `{crm}_list_deals` | List deals with pagination |
| `{crm}_get_deal` | Get deal by ID |
| `{crm}_create_deal` | Create new deal |
| `{crm}_update_deal` | Update deal |
| `{crm}_move_deal_stage` | Move deal to different stage |
| `{crm}_list_pipelines` | List available pipelines |

### Activities
| Tool | Description |
|------|-------------|
| `{crm}_log_call` | Log a phone call |
| `{crm}_log_email` | Log an email |
| `{crm}_create_task` | Create a task |
| `{crm}_list_activities` | List activities for a record |

## Authentication

### API Key (Simple)

```bash
npx wrangler secret put CRM_API_KEY
```

### OAuth 2.0 (User-authenticated)

```bash
npx wrangler secret put CRM_CLIENT_ID
npx wrangler secret put CRM_CLIENT_SECRET
npx wrangler secret put COOKIE_ENCRYPTION_KEY
npx wrangler kv namespace create "OAUTH_KV"
```

## Development

```bash
# Local development
bun run dev      # or: pnpm run dev / npm run dev

# Type checking
bun run typecheck

# Linting (Biome)
bun run lint     # Check for issues
bun run lint:fix # Auto-fix issues
bun run format   # Format code

# Deploy
bun run deploy

# Test with MCP Inspector
bunx @modelcontextprotocol/inspector
# or: pnpm dlx @modelcontextprotocol/inspector
# or: npx @modelcontextprotocol/inspector
```

## CRM-Specific Notes

When implementing for a specific CRM, refer to:

| CRM | API Version | API Docs | Auth | Rate Limits |
|-----|-------------|----------|------|-------------|
| HubSpot | v3 (stable) | [developers.hubspot.com](https://developers.hubspot.com/docs/api/overview) | OAuth 2.0 / API Key | 100/10s |
| Salesforce | v64.0 | [developer.salesforce.com](https://developer.salesforce.com/docs/apis) | OAuth 2.0 | Varies by edition |
| Dynamics 365 | v9.2 | [learn.microsoft.com](https://learn.microsoft.com/en-us/dynamics365/customerengagement/on-premises/developer/webapi/web-api-versions) | Azure AD OAuth | 60,000/5min |
| Pipedrive | **v2** | [developers.pipedrive.com](https://developers.pipedrive.com/docs/api/v2) | OAuth 2.0 / API Key | 100/10s |
| monday.com | v2 (2025-07) | [developer.monday.com](https://developer.monday.com/api-reference/docs) | API Key | 10,000 complexity/min |

> **Note:** Pipedrive API v1 is deprecated as of Dec 31, 2025. Use v2 for all new implementations.

## License

MIT
