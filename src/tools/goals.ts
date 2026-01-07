/**
 * Goal Tools
 *
 * MCP tools for goal and metric management in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

export function registerGoalTools(server: McpServer, client: CrmClient): void {
  server.tool(
    'dynamics_list_goals',
    'List goals from Dynamics 365 with pagination.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ limit, cursor, format }) => {
      try {
        const result = await client.listGoals({ limit, cursor });
        return formatResponse(result, format, 'goals');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_goal',
    'Get a single goal by ID.',
    { id: z.string(), format: z.enum(['json', 'markdown']).default('json') },
    async ({ id, format }) => {
      try {
        const goal = await client.getGoal(id);
        return formatResponse(goal, format, 'goal');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_goal',
    'Create a new goal.',
    {
      title: z.string().describe('Goal title'),
      metricId: z.string().describe('Goal metric ID'),
      goalOwnerId: z.string().describe('User or team ID who owns the goal'),
      ownerType: z.enum(['systemuser', 'team']).default('systemuser'),
      goalStartDate: z.string().describe('Goal start date (ISO format)'),
      goalEndDate: z.string().describe('Goal end date (ISO format)'),
      targetInteger: z.number().int().optional(),
      targetDecimal: z.number().optional(),
      targetMoney: z.number().optional(),
      stretchTargetInteger: z.number().int().optional(),
      stretchTargetDecimal: z.number().optional(),
      stretchTargetMoney: z.number().optional(),
      parentGoalId: z.string().optional(),
      isOverridden: z.boolean().optional(),
      isFiscalPeriodGoal: z.boolean().optional(),
      fiscalPeriod: z.number().int().optional(),
      fiscalYear: z.number().int().optional(),
    },
    async (input) => {
      try {
        const goal = await client.createGoal(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Goal created', goal }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_update_goal',
    'Update an existing goal.',
    {
      id: z.string(),
      title: z.string().optional(),
      goalStartDate: z.string().optional(),
      goalEndDate: z.string().optional(),
      targetInteger: z.number().int().optional(),
      targetDecimal: z.number().optional(),
      targetMoney: z.number().optional(),
      stretchTargetInteger: z.number().int().optional(),
      stretchTargetDecimal: z.number().optional(),
      stretchTargetMoney: z.number().optional(),
      isOverridden: z.boolean().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const goal = await client.updateGoal(id, input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Goal updated', goal }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_delete_goal',
    'Delete a goal from Dynamics 365.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.deleteGoal(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: `Goal ${id} deleted` }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_recalculate_goal',
    'Recalculate the rollup values for a goal.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.recalculateGoal(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Goal recalculated' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_list_goal_metrics',
    'List goal metrics from Dynamics 365.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
    },
    async ({ limit, cursor }) => {
      try {
        const result = await client.listGoalMetrics({ limit, cursor });
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, ...result }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_goal_metric',
    'Get a single goal metric by ID.',
    { id: z.string() },
    async ({ id }) => {
      try {
        const metric = await client.getGoalMetric(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, metric }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
