// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');

// --- searchMessageByMetadata: Primary search logic using subject and recipients ---
async function searchMessageByMetadata(client, subjectFromRequest, recipientsFromRequest, context) {
    // Helper to encode strings for OData filters and trim them.
    const encodeOData = (str) => str ? str.replace(/'/g, "''").trim() : "";

    const initialODataFilterParts = [];
    let primaryFilterCriterionUsed = ""; // To log what was used for the initial filter

    const encodedSubject = encodeOData(subjectFromRequest);
    const fullRecipientList = recipientsFromRequest 
        ? recipientsFromRequest.split(';').map(r => r.trim().toLowerCase()).filter(r => r) 
        : [];
    const firstRecipientEncoded = (fullRecipientList.length > 0) ? encodeOData(fullRecipientList[0]) : null;

    // Strategy: Prioritize subject for initial filter. If no subject, use first recipient.
    if (encodedSubject) {
        initialODataFilterParts.push(`subject eq '${encodedSubject}'`);
        primaryFilterCriterionUsed = "subject";
    } else if (firstRecipientEncoded) {
        initialODataFilterParts.push(`toRecipients/any(r: r/emailAddress/address eq '${firstRecipientEncoded}')`);
        primaryFilterCriterionUsed = "firstRecipient";
    }
    
    if (initialODataFilterParts.length === 0) {
        if (context && context.log) {
            context.log.warn("searchMessageByMetadata: Called without a valid subject or any recipients for the initial filter. Search is aborted as it would be too broad.");
        }
        throw new Error("Search requires at least a subject or recipients to be specified for filtering.");
    }

    const initialODataFilter = initialODataFilterParts.join(' and ');
    if (context && context.log) context.log.info(`searchMessageByMetadata: Constructing OData filter for Graph API using ${primaryFilterCriterionUsed}: ${initialODataFilter}`);

    try {
        // Fetch messages matching the simplified initial filter.
        // Removed .orderby() from API call. Increased .top() as we sort client-side.
        const response = await client.api('/me/messages')
            .filter(initialODataFilter)
            .top(15) // Fetch more messages to sort client-side for "latest"
            .select('id,receivedDateTime,subject,toRecipients')
            .get();

        const validatedMessages = [];

        if (response.value && response.value.length > 0) {
            if (context && context.log) context.log.info(`searchMessageByMetadata: Initial query (using ${primaryFilterCriterionUsed}) returned ${response.value.length} message(s). Performing detailed client-side validation...`);
            
            for (const message of response.value) {
                let subjectMatch = true;
                if (subjectFromRequest) {
                    const messageSubjectNormalized = message.subject ? message.subject.trim().toLowerCase() : "";
                    subjectMatch = (messageSubjectNormalized === (encodedSubject ? encodedSubject.toLowerCase() : ""));
                    if (!subjectMatch && context && context.log) {
                        // Log only if it's a mismatch for a subject we are actually checking
                        if (encodedSubject) {
                           context.log.info(`searchMessageByMetadata: Message ${message.id} (Received: ${message.receivedDateTime}): Subject mismatch. Expected: "${encodedSubject.toLowerCase()}", Actual: "${messageSubjectNormalized}".`);
                        }
                    }
                }
                if (!subjectMatch) continue;

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
                                    context.log.info(`searchMessageByMetadata: Message ${message.id} (Received: ${message.receivedDateTime}): Recipient mismatch. Expected: "${expectedRecipient}", Found in msg: [${Array.from(messageRecipientsSet).join(', ')}].`);
                                }
                                break; 
                            }
                        }
                        if (allExpectedRecipientsFound) recipientsMatch = true; 
                    } else if (context && context.log) { 
                        context.log.info(`searchMessageByMetadata: Message ${message.id} (Received: ${message.receivedDateTime}): Recipient mismatch. Expected ${fullRecipientList.length} recipients, but message has none.`);
                    }
                }
                if (!recipientsMatch) continue; 

                // If all validations pass, add to our collection for sorting
                validatedMessages.push(message);
            }

            if (validatedMessages.length > 0) {
                // Sort the validated messages by receivedDateTime in descending order (latest first)
                validatedMessages.sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime));
                
                const latestMatchingMessage = validatedMessages[0];
                if (context && context.log) context.log.info(`searchMessageByMetadata: SUCCESS - Found ${validatedMessages.length} fully matching message(s). Latest is ID: ${latestMatchingMessage.id} (Received: "${latestMatchingMessage.receivedDateTime}", Subject: "${latestMatchingMessage.subject}").`);
                return latestMatchingMessage.id;
            } else {
                 if (context && context.log) context.log.info("searchMessageByMetadata: No messages passed detailed client-side validation after initial query.");
            }

        } else if (context && context.log) { 
            context.log.info(`searchMessageByMetadata: Initial OData query returned no messages. Filter used (based on ${primaryFilterCriterionUsed}): "${initialODataFilter}"`);
        }
        return null; 
    } catch (err) {
        if (context && context.log) context.log.error(`searchMessageByMetadata: Error during Graph API call (Filter was (based on ${primaryFilterCriterionUsed}): "${initialODataFilter}"): ${err.message}`);
        throw new Error(`Error in searchMessageByMetadata: ${err.message}`); 
    }
}

// --- Main Azure Function Handler ---
// (The rest of the main handler function remains the same as in the provided immersive,
//  as the changes are localized to searchMessageByMetadata)
module.exports = async function (context, req) {
    context.log("Processing email forwarding request (Metadata Search Only Mode)...");

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
        } = req.body || {};
        
        let messageIdToProcess = null;

        context.log.info(`Attempting metadata search. Criteria: Subject="${subject}", Recipients="${recipients}"`);
        try {
            messageIdToProcess = await searchMessageByMetadata(client, subject, recipients, context);
            
            if (messageIdToProcess) {
                context.log(`Metadata search successful. Found message with Graph REST ID: "${messageIdToProcess}".`);
            } else {
                const recipientsForLog = typeof recipients === 'string' ? recipients : JSON.stringify(recipients);
                context.log.error(`Metadata search did not find a matching message. Criteria: Subject="${subject}", Recipients="${recipientsForLog}"`);
                context.res = { status: 404, body: { success: false, error: `No message found via metadata search matching criteria (Subject: "${subject}", Recipients: "${recipientsForLog}")` } };
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
