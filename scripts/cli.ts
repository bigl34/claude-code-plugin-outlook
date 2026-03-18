#!/usr/bin/env npx tsx
/**
 * Outlook Email Manager CLI
 *
 * Zod-validated CLI for Outlook/MS365 operations via MCP.
 */

import { z, createCommand, runCli, cacheCommands, cliTypes, wrapUntrustedField, buildSafeOutput } from "@local/cli-utils";
import { OutlookMCPClient } from "./mcp-client.js";

// ==================== Content Safety Helpers ====================

function wrapMessageItem(msg: any) {
  return {
    metadata: {
      id: msg.id,
      receivedDateTime: msg.receivedDateTime,
      isRead: msg.isRead,
      hasAttachments: msg.hasAttachments,
      importance: msg.importance,
      conversationId: msg.conversationId,
    },
    content: {
      subject: wrapUntrustedField("subject", msg.subject, { maxChars: 500 }),
      senderName: wrapUntrustedField("senderName", msg.from?.emailAddress?.name, { maxChars: 200 }),
      senderEmail: wrapUntrustedField("senderEmail", msg.from?.emailAddress?.address, { maxChars: 200 }),
      bodyPreview: wrapUntrustedField("bodyPreview", msg.bodyPreview, { maxChars: 500 }),
    },
  };
}

// Define commands with Zod schemas
const commands = {
  "list-tools": createCommand(
    z.object({}),
    async (_args, client: OutlookMCPClient) => {
      client.skipAuthCheck = true;
      const tools = await client.listTools();
      return tools.map((t: { name: string; description?: string }) => ({
        name: t.name,
        description: t.description,
      }));
    },
    "List all available MCP tools"
  ),

  // ==================== Authentication ====================
  "login": createCommand(
    z.object({}),
    async (_args, client: OutlookMCPClient) => {
      client.skipAuthCheck = true;
      return client.login();
    },
    "Authenticate with Microsoft"
  ),

  "verify-login": createCommand(
    z.object({}),
    async (_args, client: OutlookMCPClient) => {
      client.skipAuthCheck = true;
      return client.verifyLogin();
    },
    "Check authentication status"
  ),

  // ==================== Mail ====================
  "list-messages": createCommand(
    z.object({
      top: cliTypes.int(1, 1000).optional().describe("Max results"),
      skip: cliTypes.int(0).optional().describe("Skip results"),
      filter: z.string().optional().describe("OData filter"),
      orderBy: z.string().optional().describe("Sort order"),
      search: z.string().optional().describe("Full-text search query (KQL syntax)"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { top, skip, filter, orderBy, search } = args as {
        top?: number; skip?: number; filter?: string; orderBy?: string; search?: string;
      };
      const result = await client.listMailMessages({ top, skip, filter, orderBy, search }) as any;
      const messages = result?.value ?? [];
      const wrappedMessages = messages.map(wrapMessageItem);
      return buildSafeOutput(
        { command: "list-messages", count: messages.length },
        { messages: wrappedMessages }
      );
    },
    "List inbox messages"
  ),

  "list-folders": createCommand(
    z.object({}),
    async (_args, client: OutlookMCPClient) => client.listMailFolders(),
    "List mail folders"
  ),

  "list-folder-messages": createCommand(
    z.object({
      folderId: z.string().min(1).describe("Folder ID"),
      top: cliTypes.int(1, 1000).optional().describe("Max results"),
      skip: cliTypes.int(0).optional().describe("Skip results"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { folderId, top, skip } = args as { folderId: string; top?: number; skip?: number };
      const result = await client.listMailFolderMessages(folderId, { top, skip }) as any;
      const messages = result?.value ?? [];
      const wrappedMessages = messages.map(wrapMessageItem);
      return buildSafeOutput(
        { command: "list-folder-messages", count: messages.length },
        { messages: wrappedMessages }
      );
    },
    "List messages in a folder"
  ),

  "get-message": createCommand(
    z.object({
      id: z.string().min(1).describe("Message ID"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { id } = args as { id: string };
      const msg = await client.getMailMessage(id) as any;

      const contentFields: Record<string, unknown> = {
        subject: wrapUntrustedField("subject", msg?.subject, { maxChars: 500 }),
        senderName: wrapUntrustedField("senderName", msg?.from?.emailAddress?.name, { maxChars: 200 }),
        senderEmail: wrapUntrustedField("senderEmail", msg?.from?.emailAddress?.address, { maxChars: 200 }),
        body: wrapUntrustedField("body", msg?.body?.content, {
          maxChars: 8000,
          convertHtml: msg?.body?.contentType === "html",
        }),
      };

      // Wrap toRecipients display names
      if (Array.isArray(msg?.toRecipients)) {
        contentFields.toRecipients = msg.toRecipients.map((r: any, i: number) =>
          wrapUntrustedField(`toRecipient[${i}]`, r?.emailAddress?.name ?? r?.emailAddress?.address, { maxChars: 200 })
        );
      }

      // Wrap attachment filenames
      if (msg?.hasAttachments && Array.isArray(msg?.attachments)) {
        contentFields.attachmentNames = msg.attachments.map((a: any, i: number) =>
          wrapUntrustedField(`attachment[${i}]`, a?.name, { maxChars: 200 })
        );
      }

      return buildSafeOutput(
        {
          id: msg?.id,
          receivedDateTime: msg?.receivedDateTime,
          sentDateTime: msg?.sentDateTime,
          isRead: msg?.isRead,
          hasAttachments: msg?.hasAttachments,
          importance: msg?.importance,
          conversationId: msg?.conversationId,
        },
        contentFields
      );
    },
    "Get a specific message"
  ),

  "send-mail": createCommand(
    z.object({
      to: z.string().min(1).describe("Recipient email"),
      subject: z.string().min(1).describe("Email subject"),
      body: z.string().min(1).describe("Email body"),
      cc: z.string().optional().describe("CC recipient"),
      bodyType: z.string().optional().describe("Body type (text/html)"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { to, subject, body, cc, bodyType } = args as {
        to: string; subject: string; body: string; cc?: string; bodyType?: string;
      };
      return client.sendMail(to, subject, body, { cc, bodyType });
    },
    "Send an email"
  ),

  "create-draft": createCommand(
    z.object({
      to: z.string().min(1).describe("Recipient email"),
      subject: z.string().min(1).describe("Email subject"),
      body: z.string().min(1).describe("Email body"),
      cc: z.string().optional().describe("CC recipient"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { to, subject, body, cc } = args as {
        to: string; subject: string; body: string; cc?: string;
      };
      return client.createDraftEmail(to, subject, body, { cc });
    },
    "Create a draft email"
  ),

  "move-message": createCommand(
    z.object({
      id: z.string().min(1).describe("Message ID"),
      folderId: z.string().min(1).describe("Destination folder ID"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { id, folderId } = args as { id: string; folderId: string };
      return client.moveMailMessage(id, folderId);
    },
    "Move message to folder"
  ),

  "delete-message": createCommand(
    z.object({
      id: z.string().min(1).describe("Message ID"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { id } = args as { id: string };
      return client.deleteMailMessage(id);
    },
    "Delete a message"
  ),

  // ==================== Calendar ====================
  "list-calendars": createCommand(
    z.object({}),
    async (_args, client: OutlookMCPClient) => client.listCalendars(),
    "List calendars"
  ),

  "list-events": createCommand(
    z.object({
      calendarId: z.string().optional().describe("Calendar ID"),
      top: cliTypes.int(1, 1000).optional().describe("Max results"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { calendarId, top } = args as { calendarId?: string; top?: number };
      return client.listCalendarEvents({ calendarId, top });
    },
    "List calendar events"
  ),

  "get-event": createCommand(
    z.object({
      id: z.string().min(1).describe("Event ID"),
      calendarId: z.string().optional().describe("Calendar ID"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { id, calendarId } = args as { id: string; calendarId?: string };
      return client.getCalendarEvent(id, calendarId);
    },
    "Get a specific event"
  ),

  "get-calendar-view": createCommand(
    z.object({
      start: z.string().min(1).describe("Start date (ISO 8601)"),
      end: z.string().min(1).describe("End date (ISO 8601)"),
      calendarId: z.string().optional().describe("Calendar ID"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { start, end, calendarId } = args as { start: string; end: string; calendarId?: string };
      return client.getCalendarView(start, end, calendarId);
    },
    "Get events in date range"
  ),

  "create-event": createCommand(
    z.object({
      subject: z.string().min(1).describe("Event subject"),
      start: z.string().min(1).describe("Start date (ISO 8601)"),
      end: z.string().min(1).describe("End date (ISO 8601)"),
      body: z.string().optional().describe("Event body"),
      location: z.string().optional().describe("Event location"),
      attendees: z.string().optional().describe("Attendees (comma-separated)"),
      calendarId: z.string().optional().describe("Calendar ID"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { subject, start, end, body, location, attendees, calendarId } = args as {
        subject: string; start: string; end: string;
        body?: string; location?: string; attendees?: string; calendarId?: string;
      };
      return client.createCalendarEvent(subject, start, end, { body, location, attendees, calendarId });
    },
    "Create a calendar event"
  ),

  "update-event": createCommand(
    z.object({
      id: z.string().min(1).describe("Event ID"),
      subject: z.string().optional().describe("Event subject"),
      start: z.string().optional().describe("Start date (ISO 8601)"),
      end: z.string().optional().describe("End date (ISO 8601)"),
      body: z.string().optional().describe("Event body"),
      location: z.string().optional().describe("Event location"),
      calendarId: z.string().optional().describe("Calendar ID"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { id, subject, start, end, body, location, calendarId } = args as {
        id: string; subject?: string; start?: string; end?: string;
        body?: string; location?: string; calendarId?: string;
      };
      return client.updateCalendarEvent(id, { subject, start, end, body, location, calendarId });
    },
    "Update a calendar event"
  ),

  "delete-event": createCommand(
    z.object({
      id: z.string().min(1).describe("Event ID"),
      calendarId: z.string().optional().describe("Calendar ID"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { id, calendarId } = args as { id: string; calendarId?: string };
      return client.deleteCalendarEvent(id, calendarId);
    },
    "Delete a calendar event"
  ),

  // ==================== Contacts ====================
  "list-contacts": createCommand(
    z.object({
      top: cliTypes.int(1, 1000).optional().describe("Max results"),
      skip: cliTypes.int(0).optional().describe("Skip results"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { top, skip } = args as { top?: number; skip?: number };
      return client.listContacts({ top, skip });
    },
    "List contacts"
  ),

  // ==================== Tasks ====================
  "list-tasks": createCommand(
    z.object({
      listId: z.string().optional().describe("Task list ID"),
      top: cliTypes.int(1, 1000).optional().describe("Max results"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { listId, top } = args as { listId?: string; top?: number };
      return client.listTasks({ listId, top });
    },
    "List tasks"
  ),

  // ==================== Search ====================
  "search": createCommand(
    z.object({
      query: z.string().min(1).describe("Search query"),
      entityTypes: z.string().optional().describe("Entity types to search"),
    }),
    async (args, client: OutlookMCPClient) => {
      const { query, entityTypes } = args as { query: string; entityTypes?: string };
      const result = await client.search(query, entityTypes) as any;
      const hits = result?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
      const wrappedHits = hits.map((hit: any) => buildSafeOutput(
        {
          hitId: hit?.hitId,
          rank: hit?.rank,
        },
        {
          summary: wrapUntrustedField("summary", hit?.summary, { maxChars: 500 }),
          subject: wrapUntrustedField("subject", hit?.resource?.subject, { maxChars: 500 }),
          senderName: wrapUntrustedField("senderName", hit?.resource?.from?.emailAddress?.name, { maxChars: 200 }),
        }
      ));
      return buildSafeOutput(
        { command: "search", query, count: hits.length },
        { hits: wrappedHits }
      );
    },
    "Search across MS365"
  ),

  // Pre-built cache commands
  ...cacheCommands<OutlookMCPClient>(),
};

// Run CLI
runCli(commands, OutlookMCPClient, {
  programName: "outlook-cli",
  description: "Outlook/MS365 operations via MCP",
});
