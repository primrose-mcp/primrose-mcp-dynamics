/**
 * Environment Bindings
 *
 * Type definitions for Cloudflare Worker environment variables and bindings.
 *
 * MULTI-TENANT ARCHITECTURE:
 * This server supports multiple tenants. Tenant-specific credentials (API keys,
 * OAuth tokens, etc.) are passed via request headers, NOT stored in wrangler
 * secrets. This allows a single server instance to serve multiple customers.
 *
 * Request Headers:
 * - X-CRM-API-Key: API key for CRM authentication
 * - X-CRM-Base-URL: (Optional) Override the default CRM API base URL
 * - X-CRM-Client-ID: (Optional) OAuth client ID
 * - X-CRM-Client-Secret: (Optional) OAuth client secret
 */

// =============================================================================
// Tenant Credentials (parsed from request headers)
// =============================================================================

export interface TenantCredentials {
  /** API Key for CRM authentication (from X-CRM-API-Key header) */
  apiKey?: string;

  /** Override CRM API base URL (from X-CRM-Base-URL header) */
  baseUrl?: string;

  /** OAuth Client ID (from X-CRM-Client-ID header) */
  clientId?: string;

  /** OAuth Client Secret (from X-CRM-Client-Secret header) */
  clientSecret?: string;

  /** OAuth Access Token (from X-CRM-Access-Token header) */
  accessToken?: string;

  /** CRM-specific: Salesforce instance URL (from X-Salesforce-Instance-URL header) */
  salesforceInstanceUrl?: string;

  /** CRM-specific: Dynamics tenant ID (from X-Dynamics-Tenant-ID header) */
  dynamicsTenantId?: string;

  /** CRM-specific: Dynamics environment URL (from X-Dynamics-Environment-URL header) */
  dynamicsEnvironmentUrl?: string;
}

/**
 * Parse tenant credentials from request headers
 */
export function parseTenantCredentials(request: Request): TenantCredentials {
  const headers = request.headers;

  return {
    apiKey: headers.get('X-CRM-API-Key') || undefined,
    baseUrl: headers.get('X-CRM-Base-URL') || undefined,
    clientId: headers.get('X-CRM-Client-ID') || undefined,
    clientSecret: headers.get('X-CRM-Client-Secret') || undefined,
    accessToken: headers.get('X-CRM-Access-Token') || undefined,
    salesforceInstanceUrl: headers.get('X-Salesforce-Instance-URL') || undefined,
    dynamicsTenantId: headers.get('X-Dynamics-Tenant-ID') || undefined,
    dynamicsEnvironmentUrl: headers.get('X-Dynamics-Environment-URL') || undefined,
  };
}

/**
 * Validate that required credentials are present for Dynamics 365
 *
 * Required headers:
 * - X-Dynamics-Environment-URL: Your Dynamics environment URL
 *
 * Plus one of:
 * - X-CRM-Access-Token: Pre-obtained OAuth access token
 * OR
 * - X-Dynamics-Tenant-ID + X-CRM-Client-ID + X-CRM-Client-Secret: For OAuth client credentials flow
 */
export function validateCredentials(credentials: TenantCredentials): void {
  // Must have environment URL
  if (!credentials.dynamicsEnvironmentUrl) {
    throw new Error(
      'Missing X-Dynamics-Environment-URL header. Provide your Dynamics environment URL (e.g., https://yourorg.crm.dynamics.com)'
    );
  }

  // Must have either access token OR client credentials
  const hasAccessToken = !!credentials.accessToken;
  const hasClientCredentials =
    credentials.dynamicsTenantId && credentials.clientId && credentials.clientSecret;

  if (!hasAccessToken && !hasClientCredentials) {
    throw new Error(
      'Missing credentials. Provide either:\n' +
        '- X-CRM-Access-Token (pre-obtained OAuth token)\n' +
        'OR\n' +
        '- X-Dynamics-Tenant-ID + X-CRM-Client-ID + X-CRM-Client-Secret (for OAuth client credentials flow)'
    );
  }
}

// =============================================================================
// Environment Configuration (from wrangler.jsonc vars and bindings)
// =============================================================================

export interface Env {
  // ===========================================================================
  // Environment Variables (from wrangler.jsonc vars)
  // ===========================================================================

  /** Maximum character limit for responses */
  CHARACTER_LIMIT: string;

  /** Default page size for list operations */
  DEFAULT_PAGE_SIZE: string;

  /** Maximum page size allowed */
  MAX_PAGE_SIZE: string;

  // ===========================================================================
  // Bindings
  // ===========================================================================

  /** KV namespace for OAuth token storage */
  OAUTH_KV?: KVNamespace;

  /** Durable Object namespace for MCP sessions */
  MCP_SESSIONS?: DurableObjectNamespace;

  /** Cloudflare AI binding (optional) */
  AI?: Ai;
}

// ===========================================================================
// Helper Functions
// ===========================================================================

/**
 * Get a numeric environment value with a default
 */
export function getEnvNumber(env: Env, key: keyof Env, defaultValue: number): number {
  const value = env[key];
  if (typeof value === 'string') {
    const parsed = parseInt(value, 10);
    return Number.isNaN(parsed) ? defaultValue : parsed;
  }
  return defaultValue;
}

/**
 * Get the character limit from environment
 */
export function getCharacterLimit(env: Env): number {
  return getEnvNumber(env, 'CHARACTER_LIMIT', 50000);
}

/**
 * Get the default page size from environment
 */
export function getDefaultPageSize(env: Env): number {
  return getEnvNumber(env, 'DEFAULT_PAGE_SIZE', 20);
}

/**
 * Get the maximum page size from environment
 */
export function getMaxPageSize(env: Env): number {
  return getEnvNumber(env, 'MAX_PAGE_SIZE', 100);
}
