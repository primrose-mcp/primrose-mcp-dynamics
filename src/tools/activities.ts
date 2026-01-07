/**
 * Activity Tools
 *
 * MCP tools for activity logging (calls, emails, tasks, etc.)
 * CUSTOMIZE: Update tool names with your CRM prefix (e.g., hubspot_log_call)
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

/**
 * Register all activity-related tools
 */
export function registerActivityTools(server: McpServer, client: CrmClient): void {
  // ===========================================================================
  // List Activities
  // ===========================================================================
  server.tool(
    'dynamics_list_activities',
    `List activities (calls, emails, tasks, etc.) from the CRM.

Args:
  - recordId: Filter by associated record ID (contact, company, or deal)
  - limit: Number of activities to return (1-100, default: 20)
  - cursor: Pagination cursor
  - format: Response format

Returns:
  Paginated list of activities.`,
    {
      recordId: z.string().optional().describe('Filter by associated record ID'),
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ recordId, limit, cursor, format }) => {
      try {
        const result = await client.listActivities({ limit, cursor, recordId });
        return formatResponse(result, format, 'activities');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Log Call
  // ===========================================================================
  server.tool(
    'dynamics_log_call',
    `Log a phone call in the CRM.

Args:
  - contactId: Contact ID the call was with
  - subject: Call subject/title
  - notes: Call notes/summary
  - durationMinutes: Call duration in minutes

Returns:
  The created call activity.`,
    {
      contactId: z.string().describe('Contact ID'),
      subject: z.string().describe('Call subject/title'),
      notes: z.string().optional().describe('Call notes/summary'),
      durationMinutes: z.number().int().optional().describe('Call duration in minutes'),
    },
    async ({ contactId, subject, notes, durationMinutes }) => {
      try {
        const activity = await client.logCall(contactId, subject, notes, durationMinutes);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: 'Call logged', activity }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Log Email
  // ===========================================================================
  server.tool(
    'dynamics_log_email',
    `Log an email in the CRM.

Args:
  - contactId: Contact ID the email was with
  - subject: Email subject
  - body: Email body content
  - direction: 'sent' for outgoing, 'received' for incoming

Returns:
  The created email activity.`,
    {
      contactId: z.string().describe('Contact ID'),
      subject: z.string().describe('Email subject'),
      body: z.string().describe('Email body content'),
      direction: z
        .enum(['sent', 'received'])
        .describe("'sent' for outgoing, 'received' for incoming"),
    },
    async ({ contactId, subject, body, direction }) => {
      try {
        const activity = await client.logEmail(contactId, subject, body, direction);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: 'Email logged', activity }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Create Task
  // ===========================================================================
  server.tool(
    'dynamics_create_task',
    `Create a task in the CRM.

Args:
  - subject: Task subject/title
  - body: Task description
  - dueDate: Due date (ISO 8601)
  - contactIds: Associated contact IDs
  - companyId: Associated company ID
  - dealId: Associated deal ID

Returns:
  The created task.`,
    {
      subject: z.string().describe('Task subject/title'),
      body: z.string().optional().describe('Task description'),
      dueDate: z.string().optional().describe('Due date (ISO 8601)'),
      contactIds: z.array(z.string()).optional().describe('Associated contact IDs'),
      companyId: z.string().optional().describe('Associated company ID'),
      dealId: z.string().optional().describe('Associated deal ID'),
    },
    async (input) => {
      try {
        const activity = await client.createActivity({
          type: 'task',
          subject: input.subject,
          body: input.body,
          dueDate: input.dueDate,
          contactIds: input.contactIds,
          companyId: input.companyId,
          dealId: input.dealId,
        });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ success: true, message: 'Task created', activity }, null, 2),
            },
          ],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Create Appointment
  // ===========================================================================
  server.tool(
    'dynamics_create_appointment',
    'Create an appointment in Dynamics 365.',
    {
      subject: z.string().describe('Appointment subject'),
      description: z.string().optional().describe('Appointment description'),
      location: z.string().optional().describe('Appointment location'),
      scheduledStart: z.string().describe('Start date/time (ISO 8601)'),
      scheduledEnd: z.string().describe('End date/time (ISO 8601)'),
      regardingId: z.string().optional().describe('Regarding record ID'),
      regardingType: z.string().optional().describe('Regarding entity type (contact, account, opportunity)'),
      requiredAttendees: z.array(z.string()).optional().describe('Required attendee IDs'),
      optionalAttendees: z.array(z.string()).optional().describe('Optional attendee IDs'),
      isAllDayEvent: z.boolean().optional().default(false),
    },
    async (input) => {
      try {
        const activity = await client.createAppointment(input);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Appointment created', activity }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Get Activity
  // ===========================================================================
  server.tool(
    'dynamics_get_activity',
    'Get a single activity by ID.',
    {
      id: z.string().describe('Activity ID'),
      activityType: z.enum(['task', 'phonecall', 'email', 'appointment', 'letter', 'fax']).optional().describe('Activity type'),
      format: z.enum(['json', 'markdown']).default('json'),
    },
    async ({ id, activityType, format }) => {
      try {
        const activity = await client.getActivity(id, activityType);
        return formatResponse(activity, format, 'activity');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Update Activity
  // ===========================================================================
  server.tool(
    'dynamics_update_activity',
    'Update an existing activity.',
    {
      id: z.string().describe('Activity ID'),
      activityType: z.enum(['task', 'phonecall', 'email', 'appointment', 'letter', 'fax']).describe('Activity type'),
      subject: z.string().optional(),
      description: z.string().optional(),
      scheduledStart: z.string().optional(),
      scheduledEnd: z.string().optional(),
      priorityCode: z.number().int().optional().describe('1=Low, 2=Normal, 3=High'),
    },
    async ({ id, activityType, ...input }) => {
      try {
        const activity = await client.updateActivity(id, activityType, input);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Activity updated', activity }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Delete Activity
  // ===========================================================================
  server.tool(
    'dynamics_delete_activity',
    'Delete an activity from Dynamics 365.',
    {
      id: z.string().describe('Activity ID'),
      activityType: z.enum(['task', 'phonecall', 'email', 'appointment', 'letter', 'fax']).describe('Activity type'),
    },
    async ({ id, activityType }) => {
      try {
        await client.deleteActivity(id, activityType);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, message: `Activity ${id} deleted` }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Complete Activity
  // ===========================================================================
  server.tool(
    'dynamics_complete_activity',
    'Mark an activity as completed.',
    {
      id: z.string().describe('Activity ID'),
      activityType: z.enum(['task', 'phonecall', 'email', 'appointment', 'letter', 'fax']).describe('Activity type'),
    },
    async ({ id, activityType }) => {
      try {
        await client.completeActivity(id, activityType);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Activity marked as completed' }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Cancel Activity
  // ===========================================================================
  server.tool(
    'dynamics_cancel_activity',
    'Cancel an activity.',
    {
      id: z.string().describe('Activity ID'),
      activityType: z.enum(['task', 'phonecall', 'email', 'appointment', 'letter', 'fax']).describe('Activity type'),
    },
    async ({ id, activityType }) => {
      try {
        await client.cancelActivity(id, activityType);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Activity cancelled' }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Create Letter
  // ===========================================================================
  server.tool(
    'dynamics_create_letter',
    'Create a letter activity in Dynamics 365.',
    {
      subject: z.string().describe('Letter subject'),
      description: z.string().optional().describe('Letter content/body'),
      regardingId: z.string().optional().describe('Regarding record ID'),
      regardingType: z.string().optional().describe('Regarding entity type'),
      address: z.string().optional().describe('Recipient address'),
    },
    async (input) => {
      try {
        const activity = await client.createLetter(input);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Letter created', activity }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Create Fax
  // ===========================================================================
  server.tool(
    'dynamics_create_fax',
    'Create a fax activity in Dynamics 365.',
    {
      subject: z.string().describe('Fax subject'),
      description: z.string().optional().describe('Fax description'),
      regardingId: z.string().optional().describe('Regarding record ID'),
      regardingType: z.string().optional().describe('Regarding entity type'),
      faxNumber: z.string().optional().describe('Fax number'),
    },
    async (input) => {
      try {
        const activity = await client.createFax(input);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Fax created', activity }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // List Activities by Type
  // ===========================================================================
  server.tool(
    'dynamics_list_activities_by_type',
    'List activities filtered by type.',
    {
      activityType: z.enum(['task', 'phonecall', 'email', 'appointment', 'letter', 'fax']).describe('Activity type to list'),
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      regardingId: z.string().optional().describe('Filter by regarding record ID'),
    },
    async ({ activityType, limit, cursor, regardingId }) => {
      try {
        const result = await client.listActivitiesByType(activityType, { limit, cursor, regardingId });
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, ...result }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  // ===========================================================================
  // Get Activity Parties
  // ===========================================================================
  server.tool(
    'dynamics_get_activity_parties',
    'Get parties (participants) for an activity.',
    { activityId: z.string().describe('Activity ID') },
    async ({ activityId }) => {
      try {
        const parties = await client.getActivityParties(activityId);
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, count: parties.length, parties }, null, 2) }],
        };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
