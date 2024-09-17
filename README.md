## ACS Copilot studio sample for Phone

This repo contains a sample / proof of concept implementation for exposing a CPS bot via ACS (Phone). Please note that this is only a sample implementation and it should not be used for production use-cases.

### Set-up
- Provision an ACS resource and acquire a phone number.
- Provision an Azure AI services multi-service account and provide the ACS resource access via Managed Identities (this is used for text to speech and speech to text).
- Create an event grid subscription for the incomingmessage event, this should be a webhook subscription that points to the webhook endpoint of this sample.
- Create a .env file with the following information:

| Key      | Value      |
| ------------- | ------------- |
| ACS_CONNECTION_STRING | The ACS connection string. |
| CALLBACK_URI | Should be the base URL for your endpoint. |
| COGNITIVE_SERVICE_ENDPOINT | Grab this from the Azure AI services multi-service account. |
| DIRECTLINE_TOKEN_ENDPOINT | Grab this from copilot studio, used for connecting to CPS via directline. |
| COPILOT_STUDIO_BOT_ID | The id of your CPS bot | 


### Local testing
For easy and quick local testing, you can use Azure devtunnels to expose an endpoint for the webhook and connect to the ACS instance from your localhost.