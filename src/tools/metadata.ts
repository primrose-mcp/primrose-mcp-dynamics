/**
 * Metadata API Tools
 *
 * MCP tools for accessing Dynamics 365 metadata (entities, attributes, options).
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError } from '../utils/formatters.js';

export function registerMetadataTools(server: McpServer, client: CrmClient): void {
  server.tool(
    'dynamics_list_entities',
    'List all entities (tables) in Dynamics 365 with their metadata.',
    {
      filter: z.string().optional().describe('Filter by entity logical name (contains)'),
      includeCustom: z.boolean().optional().default(true).describe('Include custom entities'),
    },
    async ({ filter, includeCustom }) => {
      try {
        const entities = await client.listEntityMetadata({ filter, includeCustom });
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              count: entities.length,
              entities: entities.map(e => ({
                logicalName: e.logicalName,
                displayName: e.displayName,
                description: e.description,
                isCustomEntity: e.isCustomEntity,
                primaryIdAttribute: e.primaryIdAttribute,
                primaryNameAttribute: e.primaryNameAttribute,
              })),
            }, null, 2),
          }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_entity_metadata',
    'Get detailed metadata for a specific entity.',
    { entityLogicalName: z.string().describe('Entity logical name (e.g., contact, account)') },
    async ({ entityLogicalName }) => {
      try {
        const metadata = await client.getEntityMetadata(entityLogicalName);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, metadata }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_list_entity_attributes',
    'List all attributes (fields) for an entity.',
    {
      entityLogicalName: z.string().describe('Entity logical name'),
      attributeType: z.string().optional().describe('Filter by attribute type'),
    },
    async ({ entityLogicalName, attributeType }) => {
      try {
        const attributes = await client.listEntityAttributes(entityLogicalName, attributeType);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              count: attributes.length,
              attributes: attributes.map(a => ({
                logicalName: a.logicalName,
                displayName: a.displayName,
                attributeType: a.attributeType,
                isRequired: a.isRequired,
                isCustomAttribute: a.isCustomAttribute,
                description: a.description,
              })),
            }, null, 2),
          }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_attribute_metadata',
    'Get detailed metadata for a specific attribute.',
    {
      entityLogicalName: z.string().describe('Entity logical name'),
      attributeLogicalName: z.string().describe('Attribute logical name'),
    },
    async ({ entityLogicalName, attributeLogicalName }) => {
      try {
        const metadata = await client.getAttributeMetadata(entityLogicalName, attributeLogicalName);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, metadata }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_optionset_values',
    'Get the available values for an option set (picklist) field.',
    {
      entityLogicalName: z.string().describe('Entity logical name'),
      attributeLogicalName: z.string().describe('Attribute logical name'),
    },
    async ({ entityLogicalName, attributeLogicalName }) => {
      try {
        const options = await client.getOptionSetValues(entityLogicalName, attributeLogicalName);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              count: options.length,
              options,
            }, null, 2),
          }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_global_optionset',
    'Get values for a global option set.',
    { optionSetName: z.string().describe('Global option set name') },
    async ({ optionSetName }) => {
      try {
        const options = await client.getGlobalOptionSet(optionSetName);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              count: options.length,
              options,
            }, null, 2),
          }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
