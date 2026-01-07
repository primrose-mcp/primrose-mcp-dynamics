/**
 * Response Formatting Utilities
 *
 * Helpers for formatting tool responses in JSON or Markdown.
 */

import type {
  Activity,
  Company,
  Contact,
  Deal,
  Lead,
  PaginatedResponse,
  Pipeline,
  ResponseFormat,
} from '../types/entities.js';
import { CrmApiError, formatErrorForLogging } from './errors.js';

/**
 * MCP tool response type
 * Note: Index signature required for MCP SDK 1.25+ compatibility
 */
export interface ToolResponse {
  [key: string]: unknown;
  content: Array<{ type: 'text'; text: string }>;
  isError?: boolean;
}

/**
 * Format a successful response
 */
export function formatResponse(
  data: unknown,
  format: ResponseFormat,
  entityType: string
): ToolResponse {
  if (format === 'markdown') {
    return {
      content: [{ type: 'text', text: formatAsMarkdown(data, entityType) }],
    };
  }
  return {
    content: [{ type: 'text', text: JSON.stringify(data, null, 2) }],
  };
}

/**
 * Format an error response
 */
export function formatError(error: unknown): ToolResponse {
  const errorInfo = formatErrorForLogging(error);

  let message: string;
  if (error instanceof CrmApiError) {
    message = `Error: ${error.message}`;
    if (error.retryable) {
      message += ' (retryable)';
    }
  } else if (error instanceof Error) {
    message = `Error: ${error.message}`;
  } else {
    message = `Error: ${String(error)}`;
  }

  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({ error: message, details: errorInfo }, null, 2),
      },
    ],
    isError: true,
  };
}

/**
 * Format data as Markdown
 */
function formatAsMarkdown(data: unknown, entityType: string): string {
  if (isPaginatedResponse(data)) {
    return formatPaginatedAsMarkdown(data, entityType);
  }

  if (Array.isArray(data)) {
    return formatArrayAsMarkdown(data, entityType);
  }

  if (typeof data === 'object' && data !== null) {
    return formatObjectAsMarkdown(data as Record<string, unknown>, entityType);
  }

  return String(data);
}

/**
 * Type guard for paginated response
 */
function isPaginatedResponse(data: unknown): data is PaginatedResponse<unknown> {
  return (
    typeof data === 'object' &&
    data !== null &&
    'items' in data &&
    Array.isArray((data as PaginatedResponse<unknown>).items)
  );
}

/**
 * Format paginated response as Markdown
 */
function formatPaginatedAsMarkdown(data: PaginatedResponse<unknown>, entityType: string): string {
  const lines: string[] = [];

  lines.push(`## ${capitalize(entityType)}`);
  lines.push('');

  if (data.total !== undefined) {
    lines.push(`**Total:** ${data.total} | **Showing:** ${data.count}`);
  } else {
    lines.push(`**Showing:** ${data.count}`);
  }

  if (data.hasMore) {
    lines.push(`**More available:** Yes (cursor: \`${data.nextCursor}\`)`);
  }
  lines.push('');

  if (data.items.length === 0) {
    lines.push('_No items found._');
    return lines.join('\n');
  }

  // Format items based on entity type
  switch (entityType) {
    case 'contacts':
      lines.push(formatContactsTable(data.items as Contact[]));
      break;
    case 'companies':
      lines.push(formatCompaniesTable(data.items as Company[]));
      break;
    case 'deals':
      lines.push(formatDealsTable(data.items as Deal[]));
      break;
    case 'activities':
      lines.push(formatActivitiesTable(data.items as Activity[]));
      break;
    case 'leads':
      lines.push(formatLeadsTable(data.items as Lead[]));
      break;
    default:
      lines.push(formatGenericTable(data.items));
  }

  return lines.join('\n');
}

/**
 * Format contacts as Markdown table
 */
function formatContactsTable(contacts: Contact[]): string {
  const lines: string[] = [];
  lines.push('| ID | Name | Email | Phone | Title |');
  lines.push('|---|---|---|---|---|');

  for (const contact of contacts) {
    const name =
      contact.fullName || `${contact.firstName || ''} ${contact.lastName || ''}`.trim() || '-';
    lines.push(
      `| ${contact.id} | ${name} | ${contact.email || '-'} | ${contact.phone || '-'} | ${contact.title || '-'} |`
    );
  }

  return lines.join('\n');
}

/**
 * Format companies as Markdown table
 */
function formatCompaniesTable(companies: Company[]): string {
  const lines: string[] = [];
  lines.push('| ID | Name | Domain | Industry | Employees |');
  lines.push('|---|---|---|---|---|');

  for (const company of companies) {
    lines.push(
      `| ${company.id} | ${company.name} | ${company.domain || '-'} | ${company.industry || '-'} | ${company.numberOfEmployees || '-'} |`
    );
  }

  return lines.join('\n');
}

/**
 * Format deals as Markdown table
 */
function formatDealsTable(deals: Deal[]): string {
  const lines: string[] = [];
  lines.push('| ID | Name | Amount | Stage | Close Date |');
  lines.push('|---|---|---|---|---|');

  for (const deal of deals) {
    const amount = deal.amount ? `${deal.currency || '$'}${deal.amount.toLocaleString()}` : '-';
    lines.push(
      `| ${deal.id} | ${deal.name} | ${amount} | ${deal.stage || '-'} | ${deal.closeDate || '-'} |`
    );
  }

  return lines.join('\n');
}

/**
 * Format activities as Markdown table
 */
function formatActivitiesTable(activities: Activity[]): string {
  const lines: string[] = [];
  lines.push('| ID | Type | Subject | Status | Date |');
  lines.push('|---|---|---|---|---|');

  for (const activity of activities) {
    lines.push(
      `| ${activity.id} | ${activity.type} | ${activity.subject} | ${activity.status || '-'} | ${activity.activityDate || activity.createdAt || '-'} |`
    );
  }

  return lines.join('\n');
}

/**
 * Format leads as Markdown table
 */
function formatLeadsTable(leads: Lead[]): string {
  const lines: string[] = [];
  lines.push('| ID | Name | Company | Email | Status |');
  lines.push('|---|---|---|---|---|');

  for (const lead of leads) {
    const name =
      lead.fullName || `${lead.firstName || ''} ${lead.lastName || ''}`.trim() || '-';
    lines.push(
      `| ${lead.id} | ${name} | ${lead.companyName || '-'} | ${lead.email || '-'} | ${lead.status || '-'} |`
    );
  }

  return lines.join('\n');
}

/**
 * Format a generic array as Markdown table
 */
function formatGenericTable(items: unknown[]): string {
  if (items.length === 0) return '_No items_';

  const first = items[0] as Record<string, unknown>;
  const keys = Object.keys(first).slice(0, 5); // Limit columns

  const lines: string[] = [];
  lines.push(`| ${keys.join(' | ')} |`);
  lines.push(`|${keys.map(() => '---').join('|')}|`);

  for (const item of items) {
    const record = item as Record<string, unknown>;
    const values = keys.map((k) => String(record[k] ?? '-'));
    lines.push(`| ${values.join(' | ')} |`);
  }

  return lines.join('\n');
}

/**
 * Format an array as Markdown
 */
function formatArrayAsMarkdown(data: unknown[], entityType: string): string {
  if (entityType === 'pipelines') {
    return formatPipelinesAsMarkdown(data as Pipeline[]);
  }
  return formatGenericTable(data);
}

/**
 * Format pipelines as Markdown
 */
function formatPipelinesAsMarkdown(pipelines: Pipeline[]): string {
  const lines: string[] = [];

  for (const pipeline of pipelines) {
    lines.push(`### ${pipeline.name}${pipeline.isDefault ? ' *(default)*' : ''}`);
    lines.push(`**ID:** \`${pipeline.id}\``);
    lines.push('');
    lines.push('| Stage | ID | Probability |');
    lines.push('|---|---|---|');

    for (const stage of pipeline.stages) {
      const prob = stage.probability !== undefined ? `${stage.probability}%` : '-';
      lines.push(`| ${stage.name} | \`${stage.id}\` | ${prob} |`);
    }
    lines.push('');
  }

  return lines.join('\n');
}

/**
 * Format a single object as Markdown
 */
function formatObjectAsMarkdown(data: Record<string, unknown>, entityType: string): string {
  const lines: string[] = [];
  lines.push(`## ${capitalize(entityType.replace(/s$/, ''))}`);
  lines.push('');

  for (const [key, value] of Object.entries(data)) {
    if (value === null || value === undefined) continue;

    if (typeof value === 'object') {
      lines.push(`**${formatKey(key)}:**`);
      lines.push('```json');
      lines.push(JSON.stringify(value, null, 2));
      lines.push('```');
    } else {
      lines.push(`**${formatKey(key)}:** ${value}`);
    }
  }

  return lines.join('\n');
}

/**
 * Capitalize first letter
 */
function capitalize(str: string): string {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

/**
 * Format a key for display (camelCase to Title Case)
 */
function formatKey(key: string): string {
  return key
    .replace(/([A-Z])/g, ' $1')
    .replace(/^./, (str) => str.toUpperCase())
    .trim();
}
