// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');

// --- searchMessageByMetadata: Search by subject and body content snippet, then validate recipients ---
async function searchMessageByMetadata(client, subjectFromRequest, recipientsFromRequest, contentSnippetFromRequest, context) {
    // Helper to encode strings for OData filters and trim them.
    const encodeOData = (str) => str ? str.replace(/'/g, "''").trim() : "";

    const encodedSubject = encodeOData(subjectFromRequest);
    // Ensure contentSnippet is also trimmed. Graph 'contains' can be sensitive to leading/trailing whitespace.
    const encodedContentSnippet = encodeOData(contentSnippetFromRequest); 

    const initialODataFilterParts = [];
    let primaryFilterCriteriaUsed = []; // To log what was used for the initial filter

    // Add subject to filter if provided and not an empty string after trimming and encoding.
    if (encodedSubject) {
        initialODataFilterParts.push(`subject eq '${encodedSubject}'`);
        primaryFilterCriteriaUsed.push("subject");
    }
    // Add content snippet to filter if provided and not an empty string after trimming and encoding.
    if (encodedContentSnippet) {
        initialODataFilterParts.push(`contains(body/content, '${encodedContentSnippet}')`);
        primaryFilterCriteriaUsed.push("contentSnippet");
    }
    
    // If neither subject nor content snippet was provided, the search is too broad.
    if (initialODataFilterParts.length === 0) {
        if (context && context.log) {
            context.log.warn("searchMessageByMetadata: Called without subject or content snippet. Search is aborted as it would be too broad.");
        }
        throw new Error("Search requires at least a subject or a content snippet to be specified for filtering.");
    }

    const initialODataFilter = initialODataFilterParts.join(' and ');
    if (context && context.log) context.log.info(`searchMessageByMetadata: Constructing OData filter using [${primaryFilterCriteriaUsed.join(', ')}]: ${initialODataFilter}`);

    try {
        const response = await client.api('/me/messages')
            .filter(initialODataFilter)
            .top(15)
            .select('id,receivedDateTime,subject,toRecipients,bodyPreview')
            .get();

        const validatedMessages = [];

        if (response.value && response.value.length > 0) {
            if (context && context.log) context.log.info(`searchMessageByMetadata: Initial query (using [${primaryFilterCriteriaUsed.join(', ')}]) returned ${response.value.length} message(s). Performing detailed client-side validation for recipients...`);
            const fullRecipientList = recipientsFromRequest 
                ? recipientsFromRequest.split(';').map(r => r.trim().toLowerCase()).filter(r => r) 
                : [];

            for (const message of response.value) {
                let recipientsMatch = true;
                if (fullRecipientList.length > 0) {
                    recipientsMatch = false;
                    if (message.toRecipients && message.toRecipients.length > 0) {
                        const messageRecipientsSet = new Set(
                            message.toRecipients.map(r => r.emailAddress && r.emailAddress.address ? r.emailAddress.address.toLowerCase() : null).filter(Boolean)
                        );
                        let allExpectedRecipientsFound = true;
                        for (const expectedRecipient of fullRecipientList) {
                            if (!messageRecipientsSet.has(expectedRecipient)) {
                                allExpectedRecipientsFound = false;
                                if (context && context.log) {
                                    context.log.info(`searchMessageByMetadata: Message ${message.id} (Subject: "${message.subject}"): Recipient mismatch. Expected: "${expectedRecipient}".`);
                                }
                                break;
                            }
                        }
                        if (allExpectedRecipientsFound) recipientsMatch = true;
                    } else if (context && context.log) {
                        context.log.info(`searchMessageByMetadata: Message ${message.id} (Subject: "${message.subject}"): Expected recipients but none present.`);
                    }
                }
                if (recipientsMatch) {
                    validatedMessages.push(message);
                }
            }

            if (validatedMessages.length > 0) {
                validatedMessages.sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime));
                const latestMatchingMessage = validatedMessages[0];
                if (context && context.log) context.log.info(`searchMessageByMetadata: SUCCESS - Latest ID: ${latestMatchingMessage.id} (Received: ${latestMatchingMessage.receivedDateTime}).`);
                return latestMatchingMessage.id;
            } else {
                if (context && context.log) context.log.info("searchMessageByMetadata: No messages passed recipient validation.");
            }
        } else if (context && context.log) {
            context.log.info(`searchMessageByMetadata: No messages returned. Filter: "${initialODataFilter}"`);
        }
        return null;
    } catch (err) {
        if (context && context.log) context.log.error(`searchMessageByMetadata: Graph API call error: ${err.message}`);
        throw new Error(`Error in searchMessageByMetadata: ${err.message}`);
    }
}

// --- Main Azure Function Handler ---
module.exports = async function (context, req) {
    context.log("Processing email forwarding request (EWSID or Metadata Search)...");

    try {
        context.log(`Request Headers: ${JSON.stringify(req.headers)}`);
        context.log(`Request Body: ${JSON.stringify(req.body || {})}`);

        const authHeader = req.headers.authorization || '';
        if (!authHeader.startsWith('Bearer ')) {
            context.log.error("Unauthorized: No token.");
            context.res = { status: 401, body: { success: false, error: "Unauthorized: No token" } };
            return;
        }
        const accessToken = authHeader.substring(7);
        const client = getAuthenticatedClient(accessToken);
        context.log("Graph client ready.");

        const { ewsItemId, subject, recipients, contentSnippet, userEmail, useMetadataSearch } = req.body || {};
        let messageIdToProcess = null;

        if (ewsItemId && typeof ewsItemId === 'string' && ewsItemId.trim()) {
            try {
                context.log.info(`Converting EWS ID: ${ewsItemId}`);
                const translateResponse = await client.api('/me/translateExchangeIds').post({
                    inputIds: [ewsItemId], targetIdType: 'restId', sourceIdType: 'ewsId'
                });
                if (translateResponse && Array.isArray(translateResponse.value) && translateResponse.value.length > 0) {
                    messageIdToProcess = translateResponse.value[0].targetId;
                    context.log.info(`Converted to Graph ID: ${messageIdToProcess}`);
                } else {
                    context.log.error("translateExchangeIds returned no value.");
                }
            } catch (convertError) {
                context.log.error(`translateExchangeIds error: ${convertError.message}`);
                if (!useMetadataSearch) {
                    context.res = { status: 400, body: { success: false, error: convertError.message } };
                    return;
                }
                context.log.info("Falling back to metadata search.");
            }
        }

        if (!messageIdToProcess && useMetadataSearch) {
            context.log.info(`Metadata search with Subject: ${subject}`);
            try {
                messageIdToProcess = await searchMessageByMetadata(client, subject, recipients, contentSnippet, context);
                if (!messageIdToProcess) {
                    context.res = { status: 404, body: { success: false, error: "No match via metadata." } };
                    return;
                }
                context.log.info(`Metadata found Graph ID: ${messageIdToProcess}`);
            } catch (searchError) {
                context.log.error(`Metadata search error: ${searchError.message}`);
                context.res = { status: 500, body: { success: false, error: searchError.message } };
                return;
            }
        }

        if (!messageIdToProcess) {
            context.res = { status: 400, body: { success: false, error: "No valid message ID." } };
            return;
        }

        const encodedMessageId = encodeURIComponent(messageIdToProcess);
        context.log(`Processing message with ID: ${encodedMessageId}`);
        try {
            const message = await client.api(`/me/messages/${encodedMessageId}`)
                .select('subject,body,bodyPreview,toRecipients,ccRecipients,bccRecipients,from,hasAttachments,importance,isRead')
                .get();
            context.log(`Retrieved message: ${message.subject}`);

            let attachments = [];
            if (message.hasAttachments) {
                context.log("Fetching attachments...");
                const attachmentsResponse = await client.api(`/me/messages/${encodedMessageId}/attachments`).get();
                attachments = attachmentsResponse.value || [];
                context.log(`Found ${attachments.length} attachments.`);
            }

            context.log("Creating draft...");
            const newMessage = {
                subject: message.subject,
                body: { contentType: message.body.contentType, content: message.body.content },
                toRecipients: message.toRecipients || [],
                ccRecipients: message.ccRecipients || [],
                importance: message.importance || "normal"
            };
            const draftMessage = await client.api('/me/messages').post(newMessage);
            const draftId = encodeURIComponent(draftMessage.id);
            context.log(`Draft created with ID: ${draftMessage.id}`);

            if (attachments.length > 0) {
                context.log(`Adding ${attachments.length} attachments to draft...`);
                for (const attachment of attachments) {
                    context.log(`Attachment: ${attachment.name}`);
                    try {
                        const attachmentData = { "@odata.type": attachment["@odata.type"], name: attachment.name, contentType: attachment.contentType };
                        if (attachment["@odata.type"] === "#microsoft.graph.fileAttachment" && attachment.contentBytes) {
                            attachmentData.contentBytes = attachment.contentBytes;
                        } else if (attachment["@odata.type"] === "#microsoft.graph.itemAttachment" && attachment.item) {
                            attachmentData.item = attachment.item;
                        } else if (attachment["@odata.type"] === "#microsoft.graph.referenceAttachment" && attachment.sourceUrl && attachment.providerType) {
                            attachmentData.sourceUrl = attachment.sourceUrl;
                            attachmentData.providerType = attachment.providerType;
                            if (attachment.permission) attachmentData.permission = attachment.permission;
                            if (typeof attachment.isFolder === 'boolean') attachmentData.isFolder = attachment.isFolder;
                        } else {
                            context.log.warn(`Skipping unsupported attachment type for ${attachment.name}`);
                            continue;
                        }
                        await client.api(`/me/messages/${draftId}/attachments`).post(attachmentData);
                        context.log(`Added attachment: ${attachment.name}`);
                    } catch (attachError) {
                        context.log.error(`Error adding attachment: ${attachError.message}`);
                    }
                }
            }

            context.log("Sending draft...");
            await client.api(`/me/messages/${draftId}/send`).post({});
            context.log("Draft sent.");

            context.log("Moving original to deleted items...");
            await client.api(`/me/messages/${encodedMessageId}/move`).post({ destinationId: "deleteditems" });
            context.log("Moved to deleted items.");

            context.res = { status: 200, body: { success: true, message: "Forwarded and moved." } };
        } catch (processingError) {
            context.log.error(`Processing error: ${processingError.message}`);
            const errMsg = processingError.code ? `Graph API Error (${processingError.code}): ${processingError.message}` : processingError.message;
            context.res = { status: 500, body: { success: false, error: errMsg } };
        }
    } catch (error) {
        context.log.error(`Unhandled error: ${error.message}`);
        context.res = { status: 500, body: { success: false, error: `Critical error: ${error.message}` } };
    }
};

// --- Helper Functions ---
function getAuthenticatedClient(accessToken) {
    return Client.init({ authProvider: (done) => done(null, accessToken) });
}
