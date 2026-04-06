/**
 * Lark (Feishu) Channel for NanoClaw
 * Handles bidirectional messaging via Lark bot using WebSocket long-polling
 */
import * as Lark from '@larksuiteoapi/node-sdk';
import fs from 'fs';
import path from 'path';

import { registerChannel } from './registry.js';
import { Channel, OnInboundMessage, OnChatMetadata, NewMessage } from '../types.js';
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

    try {
      await this.client.im.v1.message.create(
        {
          params: {
            receive_id_type: 'chat_id',
          },
          data: {
            receive_id: chatId,
            msg_type: 'text',
            content: JSON.stringify({ text }),
          },
        },
      );

      logger.debug({ jid }, 'Lark message sent');
    } catch (err) {
      logger.error({ err, jid }, 'Failed to send Lark message');
      throw err;
    }
  }

  async setTyping(jid: string, isTyping: boolean): Promise<void> {
    // Lark doesn't have a built-in typing indicator
    logger.debug({ jid, isTyping }, 'Lark typing indicator (no-op)');
  }

  private handleMessage(data: any): void {
    try {
      const { message } = data;
      if (!message) {
        return;
      }

      const jid = `lark:${message.chat_id}`;

      // Parse message content
      let textContent = '';
      if (message.message_type === 'text') {
        try {
          const content = JSON.parse(message.content);
          textContent = content.text || '';
        } catch (e) {
          textContent = message.content || '';
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

      logger.debug({ jid, sender: newMessage.sender }, 'Lark message received');
      this.onMessage(jid, newMessage);
    } catch (err) {
      logger.error({ err, data }, 'Error processing Lark message');
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
    { appId: creds.appId ? 'present' : 'missing', appSecret: creds.appSecret ? 'present' : 'missing' },
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
