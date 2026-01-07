/**
 * Users and Teams Tools
 *
 * MCP tools for user and team management in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

export function registerUserTools(server: McpServer, client: CrmClient): void {
  // Users
  server.tool(
    'dynamics_list_users',
    'List system users from Dynamics 365 with pagination.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
      activeOnly: z.boolean().optional().default(true).describe('Only return active users'),
    },
    async ({ limit, cursor, format, activeOnly }) => {
      try {
        const result = await client.listUsers({ limit, cursor, activeOnly });
        return formatResponse(result, format, 'users');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_user',
    'Get a single user by ID.',
    { id: z.string(), format: z.enum(['json', 'markdown']).default('json') },
    async ({ id, format }) => {
      try {
        const user = await client.getUser(id);
        return formatResponse(user, format, 'user');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_current_user',
    'Get the currently authenticated user.',
    { format: z.enum(['json', 'markdown']).default('json') },
    async ({ format }) => {
      try {
        const user = await client.getCurrentUser();
        return formatResponse(user, format, 'user');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_search_users',
    'Search users by name or email.',
    {
      query: z.string().describe('Search query'),
      limit: z.number().int().min(1).max(100).default(20),
    },
    async ({ query, limit }) => {
      try {
        const users = await client.searchUsers(query, limit);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, count: users.length, users }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // Teams
  server.tool(
    'dynamics_list_teams',
    'List teams from Dynamics 365 with pagination.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listTeams({ limit, cursor });
        return formatResponse(result, format, 'teams');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_team',
    'Get a single team by ID.',
    { id: z.string(), format: z.enum(['json', 'markdown']).default('json') },
    async ({ id, format }) => {
      try {
        const team = await client.getTeam(id);
        return formatResponse(team, format, 'team');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_team',
    'Create a new team.',
    {
      name: z.string().describe('Team name'),
      description: z.string().optional(),
      businessUnitId: z.string().describe('Business unit ID'),
      teamType: z.number().int().optional().describe('0=Owner, 1=Access, 2=AADSecurityGroup, 3=AADOfficeGroup'),
      administratorId: z.string().optional().describe('Team administrator user ID'),
    },
    async (input) => {
      try {
        const team = await client.createTeam(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Team created', team }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_update_team',
    'Update an existing team.',
    {
      id: z.string(),
      name: z.string().optional(),
      description: z.string().optional(),
      administratorId: z.string().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const team = await client.updateTeam(id, input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Team updated', team }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_delete_team',
    'Delete a team from Dynamics 365.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.deleteTeam(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: `Team ${id} deleted` }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_add_team_member',
    'Add a user to a team.',
    {
      teamId: z.string(),
      userId: z.string(),
    },
    async ({ teamId, userId }) => {
      try {
        await client.addTeamMember(teamId, userId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'User added to team' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_remove_team_member',
    'Remove a user from a team.',
    {
      teamId: z.string(),
      userId: z.string(),
    },
    async ({ teamId, userId }) => {
      try {
        await client.removeTeamMember(teamId, userId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'User removed from team' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_list_team_members',
    'List members of a team.',
    { teamId: z.string() },
    async ({ teamId }) => {
      try {
        const members = await client.listTeamMembers(teamId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, count: members.length, members }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // Business Units
  server.tool(
    'dynamics_list_business_units',
    'List business units from Dynamics 365.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
    },
    async ({ limit, cursor }) => {
      try {
        const result = await client.listBusinessUnits({ limit, cursor });
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, ...result }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_business_unit',
    'Get a single business unit by ID.',
    { id: z.string() },
    async ({ id }) => {
      try {
        const businessUnit = await client.getBusinessUnit(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, businessUnit }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
