/**
 * Campaign Tools
 *
 * MCP tools for marketing campaign management in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

export function registerCampaignTools(server: McpServer, client: CrmClient): void {
  server.tool(
    'dynamics_list_campaigns',
    'List marketing campaigns from Dynamics 365 with pagination.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listCampaigns({ limit, cursor });
        return formatResponse(result, format, 'campaigns');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_campaign',
    'Get a single campaign by ID.',
    { id: z.string(), format: z.enum(['json', 'markdown']).default('json') },
    async ({ id, format }) => {
      try {
        const campaign = await client.getCampaign(id);
        return formatResponse(campaign, format, 'campaign');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_campaign',
    'Create a new marketing campaign.',
    {
      name: z.string().describe('Campaign name'),
      description: z.string().optional(),
      codeName: z.string().optional(),
      typeCode: z.number().int().optional().describe('1=Advertisement, 2=DirectMarketing, 3=Event, 4=CoMarketing, 5=Other'),
      proposedStart: z.string().optional(),
      proposedEnd: z.string().optional(),
      actualStart: z.string().optional(),
      actualEnd: z.string().optional(),
      budgetedCost: z.number().optional(),
      expectedRevenue: z.number().optional(),
      expectedResponse: z.number().int().optional(),
      message: z.string().optional(),
      objective: z.string().optional(),
      promotionCode: z.string().optional(),
    },
    async (input) => {
      try {
        const campaign = await client.createCampaign(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Campaign created', campaign }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_update_campaign',
    'Update an existing campaign.',
    {
      id: z.string(),
      name: z.string().optional(),
      description: z.string().optional(),
      codeName: z.string().optional(),
      typeCode: z.number().int().optional(),
      proposedStart: z.string().optional(),
      proposedEnd: z.string().optional(),
      actualStart: z.string().optional(),
      actualEnd: z.string().optional(),
      budgetedCost: z.number().optional(),
      expectedRevenue: z.number().optional(),
      expectedResponse: z.number().int().optional(),
      message: z.string().optional(),
      objective: z.string().optional(),
      promotionCode: z.string().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const campaign = await client.updateCampaign(id, input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Campaign updated', campaign }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_delete_campaign',
    'Delete a campaign from Dynamics 365.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.deleteCampaign(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: `Campaign ${id} deleted` }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_list_campaign_activities',
    'List activities for a campaign.',
    { campaignId: z.string() },
    async ({ campaignId }) => {
      try {
        const activities = await client.listCampaignActivities(campaignId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, count: activities.length, activities }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_campaign_activity',
    'Create a new campaign activity.',
    {
      campaignId: z.string(),
      subject: z.string(),
      description: z.string().optional(),
      channelTypeCode: z.number().int().optional().describe('1=Phone, 2=Appointment, 3=Letter, 4=Fax, 5=Email, 6=Other'),
      typeCode: z.number().int().optional().describe('1=Research, 2=Planning, 3=Preparation, 4=Distribution'),
      scheduledStart: z.string().optional(),
      scheduledEnd: z.string().optional(),
      budgetedCost: z.number().optional(),
    },
    async (input) => {
      try {
        const activity = await client.createCampaignActivity(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Campaign activity created', activity }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_list_campaign_responses',
    'List responses for a campaign.',
    { campaignId: z.string() },
    async ({ campaignId }) => {
      try {
        const responses = await client.listCampaignResponses(campaignId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, count: responses.length, responses }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_campaign_response',
    'Create a new campaign response.',
    {
      campaignId: z.string(),
      subject: z.string(),
      description: z.string().optional(),
      channelTypeCode: z.number().int().optional().describe('1=Email, 2=Phone, 3=Fax, 4=Letter, 5=Appointment, 6=Other'),
      responseCode: z.number().int().optional().describe('1=Interested, 2=NotInterested, 3=DoNotSend, 4=Error'),
      receivedOn: z.string().optional(),
      firstName: z.string().optional(),
      lastName: z.string().optional(),
      emailAddress: z.string().optional(),
      telephone: z.string().optional(),
      companyName: z.string().optional(),
    },
    async (input) => {
      try {
        const response = await client.createCampaignResponse(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Campaign response created', response }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_add_campaign_member',
    'Add a member (contact or lead) to a campaign.',
    {
      campaignId: z.string(),
      entityType: z.enum(['contact', 'lead', 'account']),
      entityId: z.string(),
    },
    async ({ campaignId, entityType, entityId }) => {
      try {
        await client.addCampaignMember(campaignId, entityType, entityId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Member added to campaign' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_remove_campaign_member',
    'Remove a member from a campaign.',
    {
      campaignId: z.string(),
      entityType: z.enum(['contact', 'lead', 'account']),
      entityId: z.string(),
    },
    async ({ campaignId, entityType, entityId }) => {
      try {
        await client.removeCampaignMember(campaignId, entityType, entityId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Member removed from campaign' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
