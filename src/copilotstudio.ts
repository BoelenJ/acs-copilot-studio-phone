global.XMLHttpRequest = require('xhr2');
global.WebSocket = require('ws');

const { DirectLine } = require('botframework-directlinejs');
//import { DirectLine } from 'botframework-directlinejs'

type conversationDetails = {
    token: string,
    conversationId: string,
    watermark?: string,
    directLine:  typeof DirectLine
}

type tokenResponse = {
    token: string,
    expires_in: number,
    conversationId: string
}

export class CopilotStudioConversationHelper {

    private conversations: Map<string, conversationDetails> = new Map<string, conversationDetails>();
    private tokenEndpoint: string;

    constructor(tokenEndpoint: string) {
        this.tokenEndpoint = tokenEndpoint;

    }

    public async getOrCreateConversation(conversationId: string) {

        if (!this.conversations.has(conversationId)) {

            // Get token.
            const token = await this.getToken();

            const directline = new DirectLine({
                token: token.token
            });

            const conversation: conversationDetails = {
                token: token.token,
                conversationId: token.conversationId,
                directLine: directline
            };

            this.conversations.set(conversationId, conversation);

            return conversation;
        };

        const conversation = this.conversations.get(conversationId);
        if(!conversation) return;
        // Refresh token.
        const directline = new DirectLine({
            token: conversation.token,
            conversationId: conversation.conversationId
        });
        conversation.directLine = directline;

        return conversation;
    }

    private async getToken() {
        // Get token from token endpoint.
        const response = await fetch(this.tokenEndpoint, { method: 'GET' });
        const json = await response.json();
        return json as tokenResponse;
    }
}