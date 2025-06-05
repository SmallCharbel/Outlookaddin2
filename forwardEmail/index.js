// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');

// --- searchMessageByMetadata: Search by exact receivedTime and subject, then validate recipients ---
async function searchMessageByMetadata(client, subjectFromRequest, recipientsFromRequest, receivedTimeFromRequest, context) {
    // Helper to encode strings for OData filters and trim them.
    const encodeOData = (str) => str ? str.replace(/'/g, "''").trim() : "";

    const encodedSubject = encodeOData(subjectFromRequest);

    // receivedTimeFromRequest and encodedSubject are critical for this precise search strategy.
    if (!receivedTimeFromRequest || typeof receivedTimeFromRequest !== 'string' || receivedTimeFromRequest.trim() === '') {
        if (context && context.log) {
            context.log.error("searchMessageByMetadata: receivedTimeFromRequest is missing, not a string, or empty. It is mandatory.");
        }
        throw new Error("receivedTimeFromRequest must be a valid ISO date string and is required.");
    }
    if (!encodedSubject) {
        if (context && context.log) {
            context.log.error("searchMessageByMetadata: subjectFromRequest is missing or empty after encoding. It is mandatory.");
        }
        throw new Error("subjectFromRequest is required and cannot be empty for this search strategy.");
    }

    // OData filter using exact receivedDateTime and subject.
    const initialODataFilter = `receivedDateTime eq ${encodeOData(receivedTimeFromRequest)} and subject eq '${encodedSubject}'`;

    if (context && context.log) context.log.info(`searchMessageByMetadata: Constructing OData filter for Graph API: ${initialODataFilter}`);

    try {
        // Fetch messages matching the exact time and subject.
        // No .orderby() from API call. .top(5) as it's expected to be very specific.
        const response = await client.api('/me/messages')
            .filter(initialODataFilter)
            .top(5) // Expect few (ideally 1) matches for exact time & subject.
            .select('id,receivedDateTime,subject,toRecipients')
            .get();

        const validatedMessages = []; // To store messages that pass recipient validation

        if (response.value && response.value.length > 0) {
            if (context && context.log) context.log.info(`searchMessageByMetadata: Initial query (time & subject) returned ${response.value.length} message(s). Performing recipient validation if needed...`);
            
            const fullRecipientList = recipientsFromRequest 
                ? recipientsFromRequest.split(';').map(r => r.trim().toLowerCase()).filter(r => r) 
                : [];

            for (const message of response.value) {
                // Subject and Time are already matched by the API query.
                // Now, validate recipients if they were provided in the request.
                let recipientsMatch = true; // Assume match if no recipients were specified in the request.
                
                if (fullRecipientList.length > 0) { // Only validate if recipients were actually expected.
                    recipientsMatch = false; // Reset to false, must be proven true.
                    if (message.toRecipients && message.toRecipients.length > 0) {
                        const messageRecipientsSet = new Set(
                            message.toRecipients.map(r => r.emailAddress && r.emailAddress.address ? r.emailAddress.address.toLowerCase() : null).filter(Boolean)
                        );
                        
                        let allExpectedRecipientsFound = true;
                        for (const expectedRecipient of fullRecipientList) {
                            if (!messageRecipientsSet.has(expectedRecipient)) {
                                allExpectedRecipientsFound = false; 
                                if (context && context.log) {
                                    context.log.info(`searchMessageByMetadata: Message ${message.id} (Received: ${message.receivedDateTime}, Subject: "${message.subject}"): Recipient mismatch. Expected: "${expectedRecipient}", Found in msg: [${Array.from(messageRecipientsSet).join(', ')}].`);
                                }
                                break; 
                            }
                        }
                        if (allExpectedRecipientsFound) recipientsMatch = true; 
                    } else if (context && context.log) { 
                        context.log.info(`searchMessageByMetadata: Message ${message.id} (Received: ${message.receivedDateTime}, Subject: "${message.subject}"): Recipient mismatch. Expected ${fullRecipientList.length} recipients, but message has none.`);
                    }
                }
                
                if (recipientsMatch) {
                    // If recipient validation passes (or wasn't needed), add to our collection.
                    validatedMessages.push(message);
                }
            }

            if (validatedMessages.length > 0) {
                // If multiple messages passed (e.g. exact same email sent to different recipient groups at same microsecond,
                // and recipient validation was not strict enough or not performed), sort by receivedDateTime.
                // This sort is mostly a safeguard; with exact time filter, usually 0 or 1 item is expected after recipient validation.
                validatedMessages.sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime));
                
                const latestMatchingMessage = validatedMessages[0];
                if (context && context.log) context.log.info(`searchMessageByMetadata: SUCCESS - Found ${validatedMessages.length} fully matching message(s) after all validations. Latest is ID: ${latestMatchingMessage.id} (Received: "${latestMatchingMessage.receivedDateTime}", Subject: "${latestMatchingMessage.subject}").`);
                return latestMatchingMessage.id;
            } else {
                 if (context && context.log) context.log.info("searchMessageByMetadata: No messages passed recipient validation after initial query by time & subject.");
            }

        } else if (context && context.log) { 
            context.log.info(`searchMessageByMetadata: Initial OData query returned no messages. Filter used: "${initialODataFilter}"`);
        }
        return null; // No message found matching all criteria
    } catch (err) {
        if (context && context.log) context.log.error(`searchMessageByMetadata: Error during Graph API call (Filter was: "${initialODataFilter}"): ${err.message}`);
        // Check if the error is the "too complex" one.
        if (err.message && err.message.toLowerCase().includes("restriction or sort order is too complex")) {
             context.log.error("searchMessageByMetadata: Encountered 'restriction or sort order is too complex' error even with simplified filter. This might indicate an issue with the specific subject/time values or a deeper Graph API limitation for this query pattern.");
        }
        throw new Error(`Error in searchMessageByMetadata: ${err.message}`); 
    }
}

// --- Main Azure Function Handler ---
module.exports = async function (context, req) {
    context.log("Processing email forwarding request (Filter by Time & Subject, Validate Recipients Client-Side)...");

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
        
        const {
            subject,
            recipients, 
            receivedTime, // This is now crucial for the search
        } = req.body || {};
        
        let messageIdToProcess = null;

        context.log.info(`Attempting metadata search. Criteria: Subject="${subject}", Recipients="${recipients}", ReceivedTime="${receivedTime}"`);
        try {
            // Pass receivedTime to the search function
            messageIdToProcess = await searchMessageByMetadata(client, subject, recipients, receivedTime, context);
            
            if (messageIdToProcess) {
                context.log(`Metadata search successful. Found message with Graph REST ID: "${messageIdToProcess}".`);
            } else {
                const recipientsForLog = typeof recipients === 'string' ? recipients : JSON.stringify(recipients);
                context.log.error(`Metadata search did not find a matching message. Criteria: Subject="${subject}", Recipients="${recipientsForLog}", ReceivedTime="${receivedTime}"`);
                context.res = { status: 404, body: { success: false, error: `No message found via metadata search matching criteria (Subject: "${subject}", Recipients: "${recipientsForLog}", ReceivedTime: "${receivedTime}")` } };
                return;
            }
        } catch (searchError) {
            context.log.error(`Error during metadata search: ${searchError.message}`);
            context.res = { status: 500, body: { success: false, error: `Error searching for message via metadata: ${searchError.message}` } };
            return;
        }
        
        context.log(`Proceeding to process message with Graph REST ID: "${messageIdToProcess}"`);
        
        try {
            const message = await client.api(`/me/messages/${messageIdToProcess}`)
                .select('subject,body,toRecipients,ccRecipients,bccRecipients,from,hasAttachments,importance,isRead')
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
                context.log(`Adding ${attachments.length} attachments to the new draft...`);
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

            context.log(`Sending the new message (draft ID: ${draftMessage.id})...`);
            await client.api(`/me/messages/${draftMessage.id}/send`).post({});
            context.log(`Successfully sent forwarded message. Original message ID was ${messageIdToProcess}.`);

            context.log(`Moving original message (ID: ${messageIdToProcess}) to deleted items...`);
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
