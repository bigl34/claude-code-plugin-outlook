/**
 * Outlook MCP Client
 *
 * Wrapper client for Microsoft Outlook operations via MCP server.
 * Provides access to personal Outlook account (YOUR_PERSONAL_EMAIL).
 *
 * Key features:
 * - Email: list, read, send, draft, move, delete
 * - Calendar: events, views, CRUD operations
 * - Contacts: list contacts
 * - Tasks: list tasks
 * - Search: cross-entity search
 *
 * Uses Microsoft Graph API via MCP server for all operations.
 */

import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { readFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { PluginCache, TTL, createCacheKey } from "@local/plugin-cache";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

interface MCPConfig {
  mcpServer: {
    command: string;
    args: string[];
    env?: Record<string, string>;
  };
}

// Initialize cache with namespace
const cache = new PluginCache({
  namespace: "outlook-email-manager",
  defaultTTL: TTL.FIVE_MINUTES,
});

/** Detects auth-related errors from MCP server responses (Microsoft Graph + generic HTTP). */
function isAuthError(text: string): boolean {
  const lower = text.toLowerCase();
  return (
    lower.includes("no valid token") ||
    lower.includes("token expired") ||
    lower.includes("unauthorized") ||
    lower.includes("401") ||
    lower.includes("invalid_grant") ||
    lower.includes("interaction_required") ||
    lower.includes("login required") ||
    lower.includes("unauthenticated") ||
    lower.includes("invalidauthenticationtoken") ||
    lower.includes("lifetimevalidationfailed") ||
    lower.includes("compacttoken") ||
    lower.includes("authorization_identitynotfound") ||
    lower.includes("silent token acquisition failed")
  );
}

export class OutlookMCPClient {
  private client: Client | null = null;
  private transport: StdioClientTransport | null = null;
  private config: MCPConfig;
  private connected: boolean = false;
  private connectPromise: Promise<void> | null = null;
  private cacheDisabled: boolean = false;

  /** Set to true before calling auth commands (login, verify-login, list-tools) to skip pre-flight auth check. */
  public skipAuthCheck = false;

  constructor() {
    // When compiled, __dirname is dist/, so look in parent for config.json
    const configPath = join(__dirname, "..", "config.json");
    this.config = JSON.parse(readFileSync(configPath, "utf-8"));
  }

  // ============================================
  // CACHE CONTROL
  // ============================================

  /** Disables caching for all subsequent requests. */
  disableCache(): void {
    this.cacheDisabled = true;
    cache.disable();
  }

  /** Re-enables caching after it was disabled. */
  enableCache(): void {
    this.cacheDisabled = false;
    cache.enable();
  }

  /** Returns cache statistics including hit/miss counts. */
  getCacheStats() {
    return cache.getStats();
  }

  /** Clears all cached data. @returns Number of cache entries cleared */
  clearCache(): number {
    return cache.clear();
  }

  /** Invalidates a specific cache entry by key. */
  invalidateCacheKey(key: string): boolean {
    return cache.invalidate(key);
  }

  // ============================================
  // CONNECTION MANAGEMENT
  // ============================================

  /** Establishes connection to the MCP server with pre-flight auth check. */
  async connect(): Promise<void> {
    if (this.connectPromise) return this.connectPromise;
    this.connectPromise = this.doConnect().catch(err => {
      this.connectPromise = null; // Reset so next attempt can retry after re-auth
      throw err;
    });
    return this.connectPromise;
  }

  private async doConnect(): Promise<void> {
    const env = {
      ...process.env,
      ...this.config.mcpServer.env,
    };

    this.transport = new StdioClientTransport({
      command: this.config.mcpServer.command,
      args: this.config.mcpServer.args,
      env: env as Record<string, string>,
    });

    this.client = new Client(
      { name: "outlook-cli", version: "1.0.0" },
      { capabilities: {} }
    );

    await this.client.connect(this.transport);
    this.connected = true;

    // Pre-flight auth check (skipped for login/verify-login/list-tools)
    if (!this.skipAuthCheck) {
      await this.checkAuth();
    }
  }

  /**
   * Verifies OAuth token validity by calling verify-login directly on the MCP SDK.
   * Uses this.client!.callTool() (raw SDK) — NOT this.callTool() which would recurse via connect().
   */
  private async checkAuth(): Promise<void> {
    const result = await this.client!.callTool({ name: "verify-login", arguments: {} });
    const content = result.content as Array<{ type: string; text?: string }>;
    const text = content.find(c => c.type === "text")?.text || "";

    let authenticated = false;
    try {
      const parsed = JSON.parse(text);
      authenticated = parsed?.success === true || parsed?.authenticated === true || parsed?.status === "authenticated";
    } catch {
      authenticated = text.toLowerCase().includes("authenticated") || text.toLowerCase().includes("login successful");
    }

    if (!authenticated) {
      throw new Error(
        "AUTHENTICATION_REQUIRED: Outlook OAuth token is expired or missing. " +
        "Do NOT retry this operation. " +
        "Ask the user to re-authenticate their personal Outlook account (YOUR_PERSONAL_EMAIL) " +
        "by running the login command in a terminal with browser access."
      );
    }
  }

  /** Disconnects from the MCP server. */
  async disconnect(): Promise<void> {
    if (this.client && this.connected) {
      await this.client.close();
      this.connected = false;
      this.connectPromise = null;
    }
  }

  // ============================================
  // MCP TOOLS
  // ============================================

  /** Lists available MCP tools. @returns Array of tool definitions */
  async listTools(): Promise<any[]> {
    await this.connect();
    const result = await this.client!.listTools();
    return result.tools;
  }

  /** Calls an MCP tool with arguments. */
  async callTool(name: string, args: Record<string, any> = {}): Promise<any> {
    await this.connect();

    const result = await this.client!.callTool({ name, arguments: args });
    const content = result.content as Array<{ type: string; text?: string }>;

    if (result.isError) {
      const errorText = content.find((c) => c.type === "text")?.text || "Tool call failed";
      if (isAuthError(errorText)) {
        throw new Error(
          "AUTHENTICATION_REQUIRED: " + errorText +
          " Do NOT retry this operation." +
          " Ask the user to re-authenticate their personal Outlook account (YOUR_PERSONAL_EMAIL)" +
          " by running the login command in a terminal with browser access."
        );
      }
      throw new Error(errorText);
    }

    const textContent = content.find((c) => c.type === "text");
    if (textContent?.text) {
      try {
        return JSON.parse(textContent.text);
      } catch {
        return textContent.text;
      }
    }

    return content;
  }

  // ============================================
  // AUTHENTICATION
  // ============================================

  /** Initiates OAuth login flow. Opens browser for authentication. */
  async login(): Promise<any> {
    return this.callTool("login", {});
  }

  /** Verifies current login status. @returns Login state and user info */
  async verifyLogin(): Promise<any> {
    return this.callTool("verify-login", {});
  }

  // ============================================
  // MAIL OPERATIONS
  // ============================================

  /**
   * Lists mail messages from inbox.
   *
   * @param options - Query options
   * @param options.top - Max messages to return
   * @param options.skip - Messages to skip (pagination)
   * @param options.filter - OData filter expression
   * @param options.orderBy - Sort order (e.g., "receivedDateTime desc")
   * @returns Message list with metadata
   *
   * @cached TTL: 5 minutes
   */
  async listMailMessages(options?: { top?: number; skip?: number; filter?: string; orderBy?: string }): Promise<any> {
    const cacheKey = createCacheKey("mail_messages", options || {});
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = {};
        if (options?.top) args.top = options.top;
        if (options?.skip) args.skip = options.skip;
        if (options?.filter) args.filter = options.filter;
        if (options?.orderBy) args.orderBy = options.orderBy;
        return this.callTool("list-mail-messages", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Lists all mail folders. @cached TTL: 1 hour */
  async listMailFolders(): Promise<any> {
    return cache.getOrFetch(
      "mail_folders",
      () => this.callTool("list-mail-folders", {}),
      { ttl: TTL.HOUR, bypassCache: this.cacheDisabled }
    );
  }

  /**
   * Lists messages in a specific mail folder.
   *
   * @param folderId - Mail folder ID
   * @param options - Pagination options
   * @returns Messages in the folder
   *
   * @cached TTL: 5 minutes
   */
  async listMailFolderMessages(folderId: string, options?: { top?: number; skip?: number }): Promise<any> {
    const cacheKey = createCacheKey("folder_messages", { folderId, ...options });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { mailFolderId: folderId };
        if (options?.top) args.top = options.top;
        if (options?.skip) args.skip = options.skip;
        return this.callTool("list-mail-folder-messages", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Gets full content of a mail message. @cached TTL: 15 minutes */
  async getMailMessage(messageId: string): Promise<any> {
    const cacheKey = createCacheKey("mail_message", { id: messageId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get-mail-message", { messageId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /**
   * Sends an email.
   *
   * @param to - Recipient email address
   * @param subject - Email subject
   * @param body - Email body content
   * @param options - Additional options
   * @param options.cc - CC recipients
   * @param options.bodyType - Body content type ("text" or "html")
   * @returns Send confirmation
   *
   * @invalidates mail_messages/*
   */
  async sendMail(to: string, subject: string, body: string, options?: { cc?: string; bodyType?: string }): Promise<any> {
    const args: Record<string, any> = { to, subject, body };
    if (options?.cc) args.cc = options.cc;
    if (options?.bodyType) args.bodyType = options.bodyType;
    const result = await this.callTool("send-mail", args);
    cache.invalidatePattern(/^mail_messages/);
    return result;
  }

  /**
   * Creates a draft email.
   *
   * @param to - Recipient email address
   * @param subject - Email subject
   * @param body - Email body content
   * @param options - Additional options (cc)
   * @returns Created draft details
   */
  async createDraftEmail(to: string, subject: string, body: string, options?: { cc?: string }): Promise<any> {
    const args: Record<string, any> = { to, subject, body };
    if (options?.cc) args.cc = options.cc;
    return this.callTool("create-draft-email", args);
  }

  /**
   * Moves a message to a different folder.
   *
   * @param messageId - Message ID to move
   * @param destinationFolderId - Target folder ID
   * @returns Moved message details
   *
   * @invalidates mail_messages/*, folder_messages/*, mail_message/{messageId}
   */
  async moveMailMessage(messageId: string, destinationFolderId: string): Promise<any> {
    const result = await this.callTool("move-mail-message", { messageId, destinationFolderId });
    cache.invalidatePattern(/^mail_messages/);
    cache.invalidatePattern(/^folder_messages/);
    cache.invalidate(createCacheKey("mail_message", { id: messageId }));
    return result;
  }

  /**
   * Deletes a mail message (moves to trash).
   *
   * @param messageId - Message ID to delete
   * @returns Deletion confirmation
   *
   * @invalidates mail_messages/*, folder_messages/*, mail_message/{messageId}
   */
  async deleteMailMessage(messageId: string): Promise<any> {
    const result = await this.callTool("delete-mail-message", { messageId });
    cache.invalidatePattern(/^mail_messages/);
    cache.invalidatePattern(/^folder_messages/);
    cache.invalidate(createCacheKey("mail_message", { id: messageId }));
    return result;
  }

  // ============================================
  // CALENDAR OPERATIONS
  // ============================================

  /** Lists all calendars. @cached TTL: 1 hour */
  async listCalendars(): Promise<any> {
    return cache.getOrFetch(
      "calendars",
      () => this.callTool("list-calendars", {}),
      { ttl: TTL.HOUR, bypassCache: this.cacheDisabled }
    );
  }

  /**
   * Lists calendar events.
   *
   * @param options - Query options
   * @param options.calendarId - Calendar ID (default: primary)
   * @param options.top - Max events to return
   * @returns Event list
   *
   * @cached TTL: 15 minutes
   */
  async listCalendarEvents(options?: { calendarId?: string; top?: number }): Promise<any> {
    const cacheKey = createCacheKey("calendar_events", options || {});
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = {};
        if (options?.calendarId) args.calendarId = options.calendarId;
        if (options?.top) args.top = options.top;
        return this.callTool("list-calendar-events", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Gets a calendar event by ID. @cached TTL: 15 minutes */
  async getCalendarEvent(eventId: string, calendarId?: string): Promise<any> {
    const cacheKey = createCacheKey("calendar_event", { id: eventId, calendarId });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { eventId };
        if (calendarId) args.calendarId = calendarId;
        return this.callTool("get-calendar-event", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /**
   * Gets calendar view for a date range.
   *
   * @param startDateTime - Start of range (ISO 8601)
   * @param endDateTime - End of range (ISO 8601)
   * @param calendarId - Optional calendar ID
   * @returns Events in the range with expanded recurrences
   *
   * @cached TTL: 15 minutes
   */
  async getCalendarView(startDateTime: string, endDateTime: string, calendarId?: string): Promise<any> {
    const cacheKey = createCacheKey("calendar_view", { start: startDateTime, end: endDateTime, calendarId });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { startDateTime, endDateTime };
        if (calendarId) args.calendarId = calendarId;
        return this.callTool("get-calendar-view", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /**
   * Creates a calendar event.
   *
   * @param subject - Event title
   * @param start - Start time (ISO 8601)
   * @param end - End time (ISO 8601)
   * @param options - Additional options
   * @param options.body - Event description
   * @param options.location - Event location
   * @param options.attendees - Comma-separated attendee emails
   * @param options.calendarId - Calendar ID
   * @returns Created event details
   *
   * @invalidates calendar_events/*, calendar_view/*
   */
  async createCalendarEvent(subject: string, start: string, end: string, options?: { body?: string; location?: string; attendees?: string; calendarId?: string }): Promise<any> {
    const args: Record<string, any> = { subject, start, end };
    if (options?.body) args.body = options.body;
    if (options?.location) args.location = options.location;
    if (options?.attendees) args.attendees = options.attendees;
    if (options?.calendarId) args.calendarId = options.calendarId;
    const result = await this.callTool("create-calendar-event", args);
    cache.invalidatePattern(/^calendar_events/);
    cache.invalidatePattern(/^calendar_view/);
    return result;
  }

  /**
   * Updates a calendar event.
   *
   * @param eventId - Event ID to update
   * @param options - Fields to update
   * @returns Updated event details
   *
   * @invalidates calendar_events/*, calendar_view/*, calendar_event/{eventId}
   */
  async updateCalendarEvent(eventId: string, options: { subject?: string; start?: string; end?: string; body?: string; location?: string; calendarId?: string }): Promise<any> {
    const args: Record<string, any> = { eventId, ...options };
    const result = await this.callTool("update-calendar-event", args);
    cache.invalidatePattern(/^calendar_events/);
    cache.invalidatePattern(/^calendar_view/);
    cache.invalidate(createCacheKey("calendar_event", { id: eventId, calendarId: options.calendarId }));
    return result;
  }

  /**
   * Deletes a calendar event.
   *
   * @param eventId - Event ID to delete
   * @param calendarId - Optional calendar ID
   * @returns Deletion confirmation
   *
   * @invalidates calendar_events/*, calendar_view/*, calendar_event/{eventId}
   */
  async deleteCalendarEvent(eventId: string, calendarId?: string): Promise<any> {
    const args: Record<string, any> = { eventId };
    if (calendarId) args.calendarId = calendarId;
    const result = await this.callTool("delete-calendar-event", args);
    cache.invalidatePattern(/^calendar_events/);
    cache.invalidatePattern(/^calendar_view/);
    cache.invalidate(createCacheKey("calendar_event", { id: eventId, calendarId }));
    return result;
  }

  // ============================================
  // CONTACTS
  // ============================================

  /**
   * Lists Outlook contacts.
   *
   * @param options - Pagination options
   * @param options.top - Max contacts to return
   * @param options.skip - Contacts to skip
   * @returns Contact list
   *
   * @cached TTL: 15 minutes
   */
  async listContacts(options?: { top?: number; skip?: number }): Promise<any> {
    const cacheKey = createCacheKey("contacts", options || {});
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = {};
        if (options?.top) args.top = options.top;
        if (options?.skip) args.skip = options.skip;
        return this.callTool("list-outlook-contacts", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  // ============================================
  // TASKS
  // ============================================

  /**
   * Lists Outlook tasks.
   *
   * @param options - Query options
   * @param options.listId - Task list ID
   * @param options.top - Max tasks to return
   * @returns Task list
   *
   * @cached TTL: 5 minutes
   */
  async listTasks(options?: { listId?: string; top?: number }): Promise<any> {
    const cacheKey = createCacheKey("tasks", options || {});
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = {};
        if (options?.listId) args.listId = options.listId;
        if (options?.top) args.top = options.top;
        return this.callTool("list-tasks", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  // ============================================
  // SEARCH
  // ============================================

  /**
   * Searches across Outlook entities.
   *
   * @param query - Search query
   * @param entityTypes - Comma-separated entity types (message, event, contact)
   * @returns Search results grouped by entity type
   *
   * @cached TTL: 5 minutes
   */
  async search(query: string, entityTypes?: string): Promise<any> {
    const cacheKey = createCacheKey("search", { query, entityTypes });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { query };
        if (entityTypes) args.entityTypes = entityTypes;
        return this.callTool("search", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }
}

export default OutlookMCPClient;
