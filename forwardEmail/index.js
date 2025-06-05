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
        // Using 'contains' on body/content. Note: This can be less performant or more complex for Graph API
        // than searching other indexed fields. Using a short, distinctive snippet is key.
        // Ensure the snippet doesn't contain characters that would break the OData query if not handled by encodeOData.
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
        // Fetch messages matching the initial filter. No .orderby() from API call.
        // Increased .top() to allow for more candidates for client-side sorting and validation.
        const response = await client.api('/me/messages')
            .filter(initialODataFilter)
            .top(15) // Fetch a reasonable number, e.g., 15, for client-side processing.
            .select('id,receivedDateTime,subject,toRecipients,bodyPreview') // bodyPreview might be useful for logging/verification.
            .get();

        const validatedMessages = []; // To store messages that pass recipient validation.

        if (response.value && response.value.length > 0) {
            if (context && context.log) context.log.info(`searchMessageByMetadata: Initial query (using [${primaryFilterCriteriaUsed.join(', ')}]) returned ${response.value.length} message(s). Performing detailed client-side validation for recipients...`);
            
            // Process recipients string from request into a clean list for validation.
            const fullRecipientList = recipientsFromRequest 
                ? recipientsFromRequest.split(';').map(r => r.trim().toLowerCase()).filter(r => r) 
                : [];

            for (const message of response.value) {
                // The initial API filter has already handled subject and content snippet matching.
                // Now, validate recipients if they were provided in the request.
                let recipientsMatch = true; // Assume match if no recipients were specified in the request for validation.
                
                if (fullRecipientList.length > 0) { // Only perform recipient validation if recipients were actually expected.
                    recipientsMatch = false; // Reset to false, must be proven true by finding all expected recipients.
                    if (message.toRecipients && message.toRecipients.length > 0) {
                        // Create a Set of the current message's recipients for efficient lookup.
                        const messageRecipientsSet = new Set(
                            message.toRecipients.map(r => r.emailAddress && r.emailAddress.address ? r.emailAddress.address.toLowerCase() : null).filter(Boolean)
                        );
                        
                        let allExpectedRecipientsFound = true;
                        for (const expectedRecipient of fullRecipientList) { // fullRecipientList is already lowercased.
                            if (!messageRecipientsSet.has(expectedRecipient)) {
                                allExpectedRecipientsFound = false; // An expected recipient is missing from this message.
                                if (context && context.log) {
                                    context.log.info(`searchMessageByMetadata: Message ${message.id} (Subject: "${message.subject}"): Recipient mismatch during validation. Expected: "${expectedRecipient}", Found in message: [${Array.from(messageRecipientsSet).join(', ')}].`);
                                }
                                break; // No need to check further recipients for this message.
                            }
                        }
                        if (allExpectedRecipientsFound) recipientsMatch = true; // All expected recipients were found in this message.
                    } else if (context && context.log) { // Message has no recipients, but we expected some.
                        context.log.info(`searchMessageByMetadata: Message ${message.id} (Subject: "${message.subject}"): Recipient mismatch. Expected ${fullRecipientList.length} recipients, but message has none.`);
                    }
                }
                
                if (recipientsMatch) {
                    // If recipient validation passes (or wasn't needed), add this message to our collection for sorting.
                    validatedMessages.push(message);
                }
            }

            if (validatedMessages.length > 0) {
                // Sort the fully validated messages by receivedDateTime in descending order (latest first).
                validatedMessages.sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime));
                
                const latestMatchingMessage = validatedMessages[0];
                if (context && context.log) context.log.info(`searchMessageByMetadata: SUCCESS - Found ${validatedMessages.length} fully matching message(s) after all validations. Latest is ID: ${latestMatchingMessage.id} (Received: "${latestMatchingMessage.receivedDateTime}", Subject: "${latestMatchingMessage.subject}").`);
                return latestMatchingMessage.id; // Return the ID of the latest, fully validated message.
            } else {
                 // No messages passed the recipient validation (if it was performed).
                 if (context && context.log) context.log.info("searchMessageByMetadata: No messages passed recipient validation after initial query by subject & content.");
            }

        } else if (context && context.log) { // Initial API query returned no messages.
            context.log.info(`searchMessageByMetadata: Initial OData query returned no messages. Filter used (based on [${primaryFilterCriteriaUsed.join(', ')}]): "${initialODataFilter}"`);
        }
        return null; // No message found matching all criteria.
    } catch (err) {
        if (context && context.log) context.log.error(`searchMessageByMetadata: Error during Graph API call (Filter was (based on [${primaryFilterCriteriaUsed.join(', ')}]): "${initialODataFilter}"): ${err.message}`);
        // Specifically log if it's the "too complex" error.
        if (err.message && err.message.toLowerCase().includes("restriction or sort order is too complex")) {
             context.log.error("searchMessageByMetadata: Encountered 'restriction or sort order is too complex' error. This suggests the combination of subject and contains(body/content) might still be too much for Graph API for this user's data, or the content snippet is problematic.");
        }
        throw new Error(`Error in searchMessageByMetadata: ${err.message}`); // Propagate the error.
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
            context.log.error("Unauthorized: No authorization token provided.");
            context.res = { status: 401, body: { success: false, error: "Unauthorized: No token provided" } };
            return;
        }
        const accessToken = authHeader.substring(7);
        const client = getAuthenticatedClient(accessToken);
        context.log("Graph client created with token.");
        
        // Destructure payload from request body.
        // Expecting ewsItemId, subject, recipients, contentSnippet, userEmail, useMetadataSearch
        const {
            ewsItemId,
            subject,
            recipients, 
            contentSnippet,
            userEmail,
            useMetadataSearch // boolean indicating fallback behavior
        } = req.body || {};
        
        let messageIdToProcess = null;

        // If an EWS Item ID is provided, convert it to a REST (Graph) ID using translateExchangeIds
        if (ewsItemId && typeof ewsItemId === 'string' && ewsItemId.trim()) {
            try {
                context.log.info(`Attempting to convert EWS ID to Graph ID: "${ewsItemId}"`);
                const translateResponse = await client.api('/me/translateExchangeIds').post({
                    inputIds: [ewsItemId],
                    targetIdType: 'restId',
                    sourceIdType: 'ewsId'
                });

                if (translateResponse && Array.isArray(translateResponse.value) && translateResponse.value.length > 0) {
                    // Use the targetId property returned by Graph
                    messageIdToProcess = translateResponse.value[0].targetId;
                    context.log.info(`Conversion successful. Graph Message ID: "${messageIdToProcess}"`);
                } else {
                    context.log.error("translateExchangeIds did not return a valid result.");
                }
            } catch (convertError) {
                context.log.error(`Error converting EWS ID to Graph ID: ${convertError.message}`);
                // If conversion fails and metadata search is allowed, fall back to metadata search.
                if (!useMetadataSearch) {
                    context.res = { status: 400, body: { success: false, error: `Failed to convert EWS ID: ${convertError.message}` } };
                    return;
                }
                context.log.info("Falling back to metadata search due to EWS ID conversion failure.");
            }
        }

        // If EWS conversion did not yield an ID and metadata search is enabled, perform metadata search.
        if (!messageIdToProcess && useMetadataSearch) {
            context.log.info(`Attempting metadata search. Criteria: Subject="${subject}", Recipients="${recipients}", ContentSnippet (length: ${contentSnippet ? contentSnippet.length : 0})`);
            try {
                messageIdToProcess = await searchMessageByMetadata(client, subject, recipients, contentSnippet, context);
                if (messageIdToProcess) {
                    context.log(`Metadata search successful. Found message with Graph REST ID: "${messageIdToProcess}".`);
                } else {
                    const recipientsForLog = typeof recipients === 'string' ? recipients : JSON.stringify(recipients);
                    context.log.error(`Metadata search did not find a matching message. Criteria: Subject="${subject}", Recipients="${recipientsForLog}", ContentSnippet provided.`);
                    context.res = { status: 404, body: { success: false, error: `No message found via metadata search matching criteria (Subject: "${subject}", Recipients: "${recipientsForLog}", ContentSnippet provided)` } };
                    return;
                }
            } catch (searchError) {
                context.log.error(`Error during metadata search: ${searchError.message}`);
                context.res = { status: 500, body: { success: false, error: `Error searching for message via metadata: ${searchError.message}` } };
                return;
            }
        }

        if (!messageIdToProcess) {
            // No ID found via EWS conversion or metadata search.
            context.res = { status: 400, body: { success: false, error: "No valid message ID to process. Please check EWS ID or search criteria." } };
            return;
        }

        // If messageIdToProcess is found (either via EWS conversion or metadata search), proceed with processing.
        context.log(`Proceeding to process message with Graph REST ID: "${messageIdToProcess}"`);
        
        try {
            const message = await client.api(`/me/messages/${messageIdToProcess}`)
                .select('subject,body,bodyPreview,toRecipients,ccRecipients,bccRecipients,from,hasAttachments,importance,isRead')
                .get();
            context.log(`Successfully retrieved original message: "${message.subject}" (ID: ${message.id})`);

            let attachments = [];
            if (message.hasAttachments) {
                context.log("Original message has attachments. Fetching attachment details...");
                const attachmentsResponse = await client.api(`/me/messages/${messageIdToProcess}/attachments`).get();
                attachments = attachmentsResponse.value || [];
                context.log(`Found ${attachments.length} attachments in the original message.`);
            }

            context.log("Creating new message draft for forwarding...");
            const newMessage = {
                subject: `${message.subject}`, 
                body: { contentType: message.body.contentType, content: message.body.content },
                toRecipients: message.toRecipients || [],
                ccRecipients: message.ccRecipients || [],
                importance: message.importance || "normal"
            };
            const draftMessage = await client.api('/me/messages').post(newMessage);
            context.log(`New draft message created with ID: ${draftMessage.id}. Subject: "${draftMessage.subject}"`);

            if (attachments.length > 0) {
                context.log(`Adding ${attachments.length} attachments to the new draft...");
                for (const attachment of attachments) {
                    context.log(`Processing attachment: "${attachment.name}" (Type: ${attachment["@odata.type"]})`);
                    try {
                        const attachmentData = {
                            "@odata.type": attachment["@odata.type"],
                            name: attachment.name,
                            contentType: attachment.contentType,
                        };
                        if (attachment["@odata.type"] === "#microsoft.graph.fileAttachment" && attachment.contentBytes) {
                            attachmentData.contentBytes = attachment.contentBytes;
                        } else if (attachment["@odata.type"] === "#microsoft.graph.itemAttachment" && attachment.item) {
                            attachmentData.item = attachment.item;
                        } else if (attachment["@odata.type"] === "#microsoft.graph.referenceAttachment" && attachment.sourceUrl && attachment.providerType) {
                            attachmentData.sourceUrl = attachment.sourceUrl;
                            attachmentData.providerType = attachment.providerType;
                            if (attachment.permission) attachmentData.permission = attachment.permission;
                            if (typeof attachment.isFolder === 'boolean') attachmentData.isFolder = attachment.isFolder;
                        } else if (attachment["@odata.type"] !== "#microsoft.graph.fileAttachment" && 
                                   attachment["@odata.type"] !== "#microsoft.graph.itemAttachment" && 
                                   attachment["@odata.type"] !== "#microsoft.graph.referenceAttachment") {
                             context.log.warn(`Unsupported attachment type or missing critical data for attachment "${attachment.name}" (Type: ${attachment["@odata.type"]}). Skipping.`);
                             continue; 
                        } else if(!attachmentData.contentBytes && !attachmentData.item && !attachmentData.sourceUrl && 
                                  (attachment["@odata.type"] === "#microsoft.graph.fileAttachment" || 
                                   attachment["@odata.type"] === "#microsoft.graph.itemAttachment" || 
                                   attachment["@odata.type"] === "#microsoft.graph.referenceAttachment") ) { 
                            context.log.warn(`Attachment "${attachment.name}" (Type: ${attachment["@odata.type"]}) is a known type but missing required data (e.g., contentBytes, item, sourceUrl). Skipping.`);
                            continue; 
                        }

                        await client.api(`/me/messages/${draftMessage.id}/attachments`).post(attachmentData);
                        context.log(`Successfully added attachment "${attachment.name}" to draft ${draftMessage.id}.`);
                    } catch (attachError) {
                        const errBody = attachError.body ? JSON.stringify(attachError.body) : 'N/A';
                        context.log.error(`Error adding attachment "${attachment.name}" to draft ${draftMessage.id}: ${attachError.message}. Error Body: ${errBody}`);
                    }
                }
            }

            context.log(`Sending the new message (draft ID: ${draftMessage.id})...");
            await client.api(`/me/messages/${draftMessage.id}/send`).post({});
            context.log(`Successfully sent forwarded message. Original message ID was ${messageIdToProcess}.`);

            context.log(`Moving original message (ID: ${messageIdToProcess}) to deleted items...");
            await client.api(`/me/messages/${messageIdToProcess}/move`).post({ destinationId: "deleteditems" });
            context.log(`Successfully moved original message (ID: ${messageIdToProcess}) to deleted items.`);

            context.log("Email forwarding process completed successfully.");
            context.res = { status: 200, body: { success: true, message: "Email forwarded and original moved to deleted items successfully." } };

        } catch (processingError) {
            context.log.error(`Error during message processing/forwarding for Graph REST ID "${messageIdToProcess}": ${processingError.message}`);
            let errMsg = `Error processing message (ID: "${messageIdToProcess}"): ${processingError.message}`;
            if (processingError.statusCode && processingError.code) { 
                errMsg = `Graph API Error (${processingError.code}) for message ID "${messageIdToProcess}": ${processingError.message}`;
            }
            context.res = { status: (processingError.statusCode === 404 ? 404 : 500), body: { success: false, error: errMsg, messageIdUsed: messageIdToProcess } };
        }
    } catch (error) { 
        context.log.error(`Unhandled error in Azure Function: ${error.message}`);
        context.res = { status: 500, body: { success: false, error: `Critical error in email forwarding process: ${error.message}` } };
    }
};

// --- Helper Functions ---
function getAuthenticatedClient(accessToken) {
    const client = Client.init({ authProvider: (done) => done(null, accessToken) });
    return client;
}
