
import { config } from 'dotenv';
import express from 'express';
import bodyParser from 'body-parser';
import { v4 as uuidv4 } from 'uuid';
import { createIdentifierFromRawId } from "@azure/communication-common";
import { AnswerCallOptions, CallAutomationClient, CallConnection, CallIntelligenceOptions, CallMediaRecognizeSpeechOrDtmfOptions, TextSource } from '@azure/communication-call-automation';
import { CopilotStudioConversationHelper } from './copilotstudio';

config();

let acsClient: CallAutomationClient;
let copilotStudioConversationHelper: CopilotStudioConversationHelper = new CopilotStudioConversationHelper(process.env.DIRECTLINE_TOKEN_ENDPOINT || "");
let callConnections: Map<string, { callerId: string, callConnection: CallConnection }> = new Map();
let messageQueue: Map<string, string[]> = new Map();
let mediaOperationStatusMap: Map<string, "PLAYING" | "AVAILABLE"> = new Map();

async function createAcsClient() {
    const connectionString = process.env.ACS_CONNECTION_STRING || "";
    acsClient = new CallAutomationClient(connectionString);
    console.log("Initialized ACS Client.");
}

// Start server with express.
const app = express();
app.use(bodyParser.json());
const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => {
    await createAcsClient();
    console.log(`Server running on port ${PORT}`);
});

// Routes.
app.get('/', (req: express.Request, res: express.Response) => {
    res.send('Hello World!');
});

// Event grid webhook.
app.post('/webhook', async (req: express.Request, res: express.Response) => {
    const event = req.body[0];
    const type = event.eventType;

    // Validation
    if (type === 'Microsoft.EventGrid.SubscriptionValidationEvent') {
        console.log('Validating webhook');
        const code = req.body[0].data.validationCode;
        res.status(200).send({ validationResponse: code });
        return;
    };

    // Handle incoming call.
    if (type === "Microsoft.Communication.IncomingCall") {
        console.log('Received incoming call');
        const callContext = event.data.incomingCallContext;
        const callerId = event.data.from.rawId;
        const uuid = uuidv4();
        const callbackUri = `${process.env.CALLBACK_URI}/callbacks/${uuid}?callerId=${callerId}`;
        const callIntelligenceOptions: CallIntelligenceOptions = { cognitiveServicesEndpoint: process.env.COGNITIVE_SERVICE_ENDPOINT };
        const answerCallOptions: AnswerCallOptions = { callIntelligenceOptions: callIntelligenceOptions };
        const answerResult = await acsClient.answerCall(callContext, callbackUri, answerCallOptions);
        callConnections.set(uuid, { callerId: callerId, callConnection: answerResult.callConnection });
        res.status(200).send();
        return;
    }
});

// Callbacks that will be triggered during the call.
app.post('/callbacks/:uuid', async (req: express.Request, res: express.Response) => {

    const event = req.body[0];
    const type = event.type;

    const uuid = req.params.uuid;
    const callDefinition = callConnections.get(uuid);

    if (!callDefinition) {
        res.status(404).send();
        return;
    };

    if (type === "Microsoft.Communication.CallConnected") {
        console.log('Call connected');

        // Start conversation with copilot studio.
        const conversation = await copilotStudioConversationHelper.getOrCreateConversation(uuid);

        if (!conversation) return;

        // Create status map for media operations.
        mediaOperationStatusMap.set(uuid, "AVAILABLE");

        conversation.directLine.activity$.filter((activity: { type: string; from: { id: string; }; }) => activity.type === 'message' && activity.from.id === process.env.COPILOT_STUDIO_BOT_ID || "").subscribe(async (activity: any) => {
            await queueOrPlayMessage(uuid, callDefinition.callConnection, callDefinition.callerId, activity.speak);
        });

        sendMessageAndWaitForReply(callDefinition.callConnection, callDefinition.callerId, "Hello, how can I help you?");
        res.status(200).send();
        return;
    }

    if (type === "Microsoft.Communication.ParticipantsUpdated") {
        console.log('Participants updated');
        res.status(200).send();
        return;
    }

    if (event.type === "Microsoft.Communication.PlayCompleted") {
        console.log('Play completed');
        checkForNextMessage(uuid, callDefinition.callConnection, callDefinition.callerId);
    }

    if (type === "Microsoft.Communication.RecognizeCompleted") {
        console.log('Recognize completed');
        if (event.data.recognitionType === "speech") {
            const text = event.data.speechResult.speech;
            console.log(`Recognized text: ${text}`);

            // Send to Copilot studio.
            const conversation = await copilotStudioConversationHelper.getOrCreateConversation(uuid);
            if (!conversation) return;
            conversation.directLine.postActivity({
                type: "message",
                from: { id: uuid },
                text: text
            }).subscribe();
            //await checkForNextMessage(uuid, callDefinition.callConnection, callDefinition.callerId);

        } else if (event.data.recognitionType === "dtmf") {
            const dtmf = event.data.dtmfResult.tones;
            const conversation = await copilotStudioConversationHelper.getOrCreateConversation(uuid);
            if (!conversation) return;
            conversation.directLine.postActivity({
                type: "message",
                from: { id: uuid },
                text: "/DTMFKey #",
                textFormat: "plain",
            }).subscribe();
            console.log(`Recognized dtmf: ${dtmf}`);
        }
        res.status(200).send();
        return;
    }
});

async function sendMessageAndWaitForReply(callConnection: CallConnection, callerId: string, message: string) {

    const play: TextSource = { text: message, voiceName: "en-US-NancyNeural", kind: "textSource" }
    const recognizeOptions: CallMediaRecognizeSpeechOrDtmfOptions = message === "" ? {
        endSilenceTimeoutInSeconds: 1,
        initialSilenceTimeoutInSeconds: 15,
        interruptPrompt: false,
        maxTonesToCollect: 1,
        kind: "callMediaRecognizeSpeechOrDtmfOptions"
    } : {
        endSilenceTimeoutInSeconds: 1,
        playPrompt: play,
        initialSilenceTimeoutInSeconds: 15,
        interruptPrompt: false,
        maxTonesToCollect: 1,
        kind: "callMediaRecognizeSpeechOrDtmfOptions"
    };

    const targetParticipant = createIdentifierFromRawId(callerId);
    await callConnection.getCallMedia().startRecognizing(targetParticipant, recognizeOptions);
}

async function playMessage(callConnection: CallConnection, callerId: string, message: string, uuid: string) {
    mediaOperationStatusMap.set(uuid, "PLAYING");
    const play: TextSource = { text: message, voiceName: "en-US-NancyNeural", kind: "textSource" }
    const targetParticipant = createIdentifierFromRawId(callerId);
    await callConnection.getCallMedia().play([play], [targetParticipant]);
    mediaOperationStatusMap.set(uuid, "AVAILABLE");

}

async function queueOrPlayMessage(uuid: string, callConnection: CallConnection, callerId: string, message: string) {


    if (!messageQueue.has(uuid)) {
        console.log("creating queue and adding first message")
        messageQueue.set(uuid, [message]);

    } else {
        console.log("adding message to queue");
        messageQueue.get(uuid)?.push(message);
    }

    await checkForNextMessage(uuid, callConnection, callerId);
}

async function checkForNextMessage(uuid: string, callConnection: CallConnection, callerId: string) {

    const messages = messageQueue.get(uuid);
    console.log("CHECKING FOR NEXT MESSAGE");
    console.log(messages);
    const mediaOperationStatus = mediaOperationStatusMap.get(uuid);
    if(!mediaOperationStatus || mediaOperationStatus == "PLAYING"){
        console.log("Media playing already, so skipping message for now.")
        return;
    }
    if (!messages) return;
    const messageLength = messages.length;
    if (messageLength > 0) {
        const message = messages.shift();
        console.log("Playing new message, not waiting for reply");
        
        await playMessage(callConnection, callerId, message || "", uuid);
    } else{
        console.log("No more message, waiting for reply");
        await sendMessageAndWaitForReply(callConnection, callerId, "");
    }
}