/**
 * Notes and Attachments Tools
 *
 * MCP tools for notes and document attachment management in Dynamics 365.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { CrmClient } from '../client.js';
import { formatError, formatResponse } from '../utils/formatters.js';

export function registerNoteTools(server: McpServer, client: CrmClient): void {
  server.tool(
    'dynamics_list_notes',
    'List notes from Dynamics 365 with pagination.',
    {
      limit: z.number().int().min(1).max(100).default(20),
      cursor: z.string().optional(),
      format: z.enum(['json', 'markdown']).default('json'),
      regardingEntityType: z.string().optional().describe('Filter by entity type (e.g., contact, account, opportunity)'),
      regardingId: z.string().optional().describe('Filter by regarding record ID'),
    },
    async ({ limit, cursor, format, regardingEntityType, regardingId }) => {
      try {
        const result = await client.listNotes({ limit, cursor, regardingEntityType, regardingId });
        return formatResponse(result, format, 'notes');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_note',
    'Get a single note by ID.',
    { id: z.string(), format: z.enum(['json', 'markdown']).default('json') },
    async ({ id, format }) => {
      try {
        const note = await client.getNote(id);
        return formatResponse(note, format, 'note');
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_create_note',
    'Create a new note attached to a record.',
    {
      subject: z.string().describe('Note subject'),
      noteText: z.string().optional().describe('Note body text'),
      regardingEntityType: z.string().describe('Entity type (e.g., contact, account, opportunity)'),
      regardingId: z.string().describe('Record ID to attach note to'),
      fileName: z.string().optional().describe('Attachment file name'),
      mimeType: z.string().optional().describe('Attachment MIME type'),
      documentBody: z.string().optional().describe('Base64 encoded file content'),
    },
    async (input) => {
      try {
        const note = await client.createNote(input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Note created', note }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_update_note',
    'Update an existing note.',
    {
      id: z.string(),
      subject: z.string().optional(),
      noteText: z.string().optional(),
      fileName: z.string().optional(),
      mimeType: z.string().optional(),
      documentBody: z.string().optional(),
    },
    async ({ id, ...input }) => {
      try {
        const note = await client.updateNote(id, input);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Note updated', note }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_delete_note',
    'Delete a note from Dynamics 365.',
    { id: z.string() },
    async ({ id }) => {
      try {
        await client.deleteNote(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: `Note ${id} deleted` }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_get_note_attachment',
    'Get the attachment content of a note (base64 encoded).',
    { id: z.string() },
    async ({ id }) => {
      try {
        const attachment = await client.getNoteAttachment(id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, ...attachment }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_list_entity_notes',
    'List all notes for a specific entity record.',
    {
      entityType: z.string().describe('Entity type (e.g., contact, account, opportunity)'),
      entityId: z.string().describe('Entity record ID'),
      limit: z.number().int().min(1).max(100).default(20),
    },
    async ({ entityType, entityId, limit }) => {
      try {
        const notes = await client.listEntityNotes(entityType, entityId, limit);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, count: notes.length, notes }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_add_attachment_to_note',
    'Add an attachment to an existing note.',
    {
      noteId: z.string(),
      fileName: z.string().describe('Attachment file name'),
      mimeType: z.string().describe('Attachment MIME type'),
      documentBody: z.string().describe('Base64 encoded file content'),
    },
    async ({ noteId, fileName, mimeType, documentBody }) => {
      try {
        await client.addAttachmentToNote(noteId, { fileName, mimeType, documentBody });
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Attachment added to note' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_remove_attachment_from_note',
    'Remove the attachment from a note.',
    { noteId: z.string() },
    async ({ noteId }) => {
      try {
        await client.removeAttachmentFromNote(noteId);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Attachment removed from note' }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );

  server.tool(
    'dynamics_search_notes',
    'Search notes by subject or content.',
    {
      query: z.string().describe('Search query'),
      limit: z.number().int().min(1).max(100).default(20),
    },
    async ({ query, limit }) => {
      try {
        const notes = await client.searchNotes(query, limit);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, count: notes.length, notes }, null, 2) }] };
      } catch (error) {
        return formatError(error);
      }
    }
  );
}
