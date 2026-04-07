/**
 * Lark (Feishu) Channel for NanoClaw
 * Handles bidirectional messaging via Lark bot using WebSocket long-polling
 */
import * as Lark from '@larksuiteoapi/node-sdk';
import Database from 'better-sqlite3';
import fs from 'fs';
import path from 'path';

import { registerChannel } from './registry.js';
import {
  Channel,
  OnInboundMessage,
  OnChatMetadata,
  NewMessage,
  FileAttachment,
} from '../types.js';
import { logger } from '../logger.js';

/**
 * Read Lark credentials from .env file
 */
function loadLarkCredentials(): { appId?: string; appSecret?: string } {
  const envFile = path.join(process.cwd(), '.env');
  let content: string;
  try {
    content = fs.readFileSync(envFile, 'utf-8');
  } catch (err) {
    logger.debug({ err }, '.env file not found');
    return {};
  }

  const result: { appId?: string; appSecret?: string } = {};

  for (const line of content.split('\n')) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eqIdx = trimmed.indexOf('=');
    if (eqIdx === -1) continue;
    const key = trimmed.slice(0, eqIdx).trim();
    let value = trimmed.slice(eqIdx + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    if (key === 'LARK_APP_ID' && value) result.appId = value;
    if (key === 'LARK_APP_SECRET' && value) result.appSecret = value;
  }

  return result;
}

export class LarkChannel implements Channel {
  name = 'lark';
  private client: Lark.Client;
  private wsClient?: Lark.WSClient;
  private onMessage: OnInboundMessage;
  private onChatMetadata: OnChatMetadata;
  private connected = false;
  private wsConnected = false;
  private appId: string;
  private appSecret: string;

  constructor(
    appId: string,
    appSecret: string,
    onMessage: OnInboundMessage,
    onChatMetadata: OnChatMetadata,
  ) {
    this.appId = appId;
    this.appSecret = appSecret;
    this.onMessage = onMessage;
    this.onChatMetadata = onChatMetadata;

    this.client = new Lark.Client({
      appId,
      appSecret,
    });
  }

  async connect(): Promise<void> {
    if (this.connected) {
      return;
    }

    logger.info('Starting Lark channel with WebSocket long-polling...');

    try {
      // Create WSClient for receiving events
      this.wsClient = new Lark.WSClient({
        appId: this.appId,
        appSecret: this.appSecret,
      });

      // Register event handler for receiving messages
      const eventDispatcher = new Lark.EventDispatcher({}).register({
        'im.message.receive_v1': async (data: any) => {
          this.handleMessage(data);
        },
      });

      // Start WebSocket connection
      await this.wsClient.start({
        eventDispatcher,
      });

      this.wsConnected = true;

      // Emit metadata for the bot's own chat
      this.onChatMetadata(
        'lark:bot_self',
        new Date().toISOString(),
        'Lark Bot',
        'lark',
        false,
      );

      this.connected = true;
      logger.info('Lark channel connected via WebSocket');
    } catch (err) {
      logger.error({ err }, 'Failed to connect Lark channel');
      throw err;
    }
  }

  async disconnect(): Promise<void> {
    if (!this.connected) {
      return;
    }

    // Stop WebSocket client by clearing the instance
    // Note: WSClient doesn't expose a stop() method, so we just clear the reference
    // The underlying WebSocket will be cleaned up by the garbage collector
    if (this.wsClient) {
      // @ts-ignore - Accessing internal wsConfig to close the connection
      if (this.wsClient.wsConfig?.wsInstance) {
        // @ts-ignore
        this.wsClient.wsConfig.wsInstance.close();
      }
      this.wsClient = undefined;
    }

    this.wsConnected = false;
    this.connected = false;
    logger.info('Lark channel disconnected');
  }

  isConnected(): boolean {
    return this.connected && this.wsConnected;
  }

  ownsJid(jid: string): boolean {
    return jid.startsWith('lark:');
  }

  async sendMessage(jid: string, text: string): Promise<void> {
    const chatId = jid.replace('lark:', '');

    // Build an interactive card (JSON 1.0) with a markdown component so
    // Claude's response renders natively (bold, lists, code blocks, etc.)
    const card = {
      config: { wide_screen_mode: true },
      header: {
        title: {
          tag: 'plain_text',
          content: 'Assistant',
        },
        template: 'blue',
      },
      elements: [
        {
          tag: 'markdown',
          content: text,
        },
      ],
    };

    try {
      await this.client.im.v1.message.create({
        params: {
          receive_id_type: 'chat_id',
        },
        data: {
          receive_id: chatId,
          msg_type: 'interactive',
          content: JSON.stringify(card),
        },
      });

      logger.debug({ jid }, 'Lark message sent');
    } catch (err) {
      logger.error({ err, jid }, 'Failed to send Lark message');
      throw err;
    }
  }

  async sendFile(jid: string, file: FileAttachment): Promise<void> {
    if (!fs.existsSync(file.filePath)) {
      throw new Error(`File not found: ${file.filePath}`);
    }

    const chatId = jid.replace('lark:', '');

    try {
      // Step 1: Upload file to Lark to get file_key
      // SDK returns responses differently per file_type — stream returns
      // { file_key } at the top level, while doc types return { data: { file_key } }
      const uploadResp = await this.client.im.v1.file.create({
        params: {
          receive_id_type: 'chat_id',
        },
        data: {
          file_type: 'stream',
          file_name: file.fileName || path.basename(file.filePath),
          file: fs.createReadStream(file.filePath),
        } as any,
      } as any);

      const fileKey =
        (uploadResp as any)?.file_key || (uploadResp as any)?.data?.file_key;
      if (!fileKey) {
        throw new Error(`File upload failed: ${JSON.stringify(uploadResp)}`);
      }

      // Step 2: Send file message using the file_key
      await this.client.im.v1.message.create({
        params: {
          receive_id_type: 'chat_id',
        },
        data: {
          receive_id: chatId,
          msg_type: 'file',
          content: JSON.stringify({ file_key: fileKey }),
        },
      });

      logger.debug({ jid, fileName: file.fileName }, 'Lark file sent');
    } catch (err) {
      logger.error({ err, jid }, 'Failed to send Lark file');
      throw err;
    }
  }

  async setTyping(jid: string, isTyping: boolean): Promise<void> {
    // Lark doesn't have a built-in typing indicator
    logger.debug({ jid, isTyping }, 'Lark typing indicator (no-op)');
  }

  private async handleMessage(data: any): Promise<void> {
    try {
      const { message } = data;
      if (!message) {
        return;
      }

      const jid = `lark:${message.chat_id}`;

      // Debug: log the raw message type and content
      logger.info(
        {
          messageType: message.message_type,
          contentPreview: (message.content || '').slice(0, 200),
        },
        'Lark raw message received',
      );

      // Parse message content
      let textContent = '';
      if (message.message_type === 'text') {
        try {
          const content = JSON.parse(message.content);
          textContent = content.text || '';
        } catch (e) {
          textContent = message.content || '';
        }
      } else if (message.message_type === 'image') {
        const imageKey = (() => {
          try {
            return JSON.parse(message.content)?.image_key || '';
          } catch {
            return '';
          }
        })();
        if (imageKey) {
          textContent = await this.downloadAndStoreFile(
            imageKey,
            'image',
            jid,
            message.sender_id?.union_id || 'unknown',
            message.message_id,
            undefined,
          );
        }
      } else if (message.message_type === 'file') {
        const fileInfo = (() => {
          try {
            return JSON.parse(message.content);
          } catch {
            return {};
          }
        })();
        const fileKey = fileInfo.file_key || '';
        const fileName = fileInfo.file_name || 'file';
        if (fileKey) {
          textContent = await this.downloadAndStoreFile(
            fileKey,
            'file',
            jid,
            message.sender_id?.union_id || 'unknown',
            message.message_id,
            fileName,
          );
        }
      }

      // Emit to main message handler
      const newMessage: NewMessage = {
        id: message.message_id,
        chat_jid: jid,
        sender: message.sender_id?.union_id || 'unknown',
        sender_name: message.sender_id?.union_id || 'Unknown User',
        content: textContent,
        timestamp: new Date().toISOString(),
        is_from_me: false,
      };

      logger.info(
        {
          jid,
          sender: newMessage.sender,
          content: newMessage.content.slice(0, 100),
        },
        'Lark message received',
      );
      this.onMessage(jid, newMessage);
    } catch (err) {
      logger.error({ err, data }, 'Error processing Lark message');
    }
  }

  private async downloadAndStoreFile(
    fileKey: string,
    fileType: 'image' | 'file',
    chatJid: string,
    sender: string,
    messageId: string,
    fileName?: string,
  ): Promise<string> {
    try {
      logger.info(
        { fileKey, fileName, messageId },
        'Lark file download starting...',
      );

      // Resolve group folder from chatJid
      const db = new Database(path.join(process.cwd(), 'store', 'messages.db'));
      const row = db
        .prepare('SELECT folder FROM registered_groups WHERE jid = ?')
        .get(chatJid) as { folder: string } | undefined;
      db.close();

      if (!row) {
        logger.error({ chatJid }, 'No registered group for this chatJid');
        return `[Received a file but couldn't find the group.]`;
      }

      // Store files in ClawWorld/ReceivedFiles/<group_folder>/
      const receivedDir = path.join(
        process.env.HOME || '/Users/zijunwu',
        'Documents',
        'ClawWorld',
        'ReceivedFiles',
        row.folder,
      );
      if (!fs.existsSync(receivedDir)) {
        fs.mkdirSync(receivedDir, { recursive: true });
      }

      // Get tenant_access_token via direct HTTP
      const tokenResp = await fetch(
        'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            app_id: this.appId,
            app_secret: this.appSecret,
          }),
        },
      );
      const tokenData = (await tokenResp.json()) as Record<string, string>;
      const token = tokenData?.tenant_access_token;
      if (!token) {
        throw new Error(
          `Failed to get Lark token: ${JSON.stringify(tokenData)}`,
        );
      }

      // Download user-sent file via messageResource.get API
      // im/v1/file/get only works for bot-uploaded files
      // For user-sent files: use messageResource.get with withTenantToken
      const downloadClient = new Lark.Client({
        appId: this.appId,
        appSecret: this.appSecret,
        disableTokenCache: true,
      });

      const safeName = fileName
        ? path.basename(fileName)
        : `${fileType}_${fileKey}.${fileType === 'image' ? 'png' : 'dat'}`;
      // Avoid overwriting existing file
      let destPath = path.join(receivedDir, safeName);
      if (fs.existsSync(destPath)) {
        const ext = path.extname(safeName);
        const base = path.basename(safeName, ext);
        let i = 1;
        while (fs.existsSync(destPath)) {
          destPath = path.join(receivedDir, `${base}_${i}${ext}`);
          i++;
        }
      }

      const resp = await downloadClient.im.v1.messageResource.get(
        {
          path: {
            message_id: messageId,
            file_key: fileKey,
          },
          params: { type: fileType === 'image' ? 'image' : 'file' },
        },
        Lark.withTenantToken(token),
      );

      await (resp as any).writeFile(destPath);

      const size = fs.statSync(destPath).size;
      logger.info(
        { fileName: safeName, destPath, size },
        'Lark file downloaded',
      );

      return `File received: **${safeName}** (${size} bytes)`;
    } catch (err) {
      logger.error({ err, fileKey }, 'Lark file download failed');
      return `[Received a file but download failed: ${err instanceof Error ? err.message : String(err)}]`;
    }
  }
}

// Self-registration
export function setupLarkChannel(opts: {
  onMessage: OnInboundMessage;
  onChatMetadata: OnChatMetadata;
}): Channel | null {
  const creds = loadLarkCredentials();

  logger.debug(
    {
      appId: creds.appId ? 'present' : 'missing',
      appSecret: creds.appSecret ? 'present' : 'missing',
    },
    'Lark channel setup',
  );

  if (!creds.appId || !creds.appSecret) {
    logger.warn(
      { appId: !!creds.appId, appSecret: !!creds.appSecret },
      'Lark channel: missing credentials in .env',
    );
    return null;
  }

  const channel = new LarkChannel(
    creds.appId,
    creds.appSecret,
    opts.onMessage,
    opts.onChatMetadata,
  );

  return channel;
}

// Auto-register when module is imported
registerChannel('lark', setupLarkChannel);
