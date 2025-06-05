// Required for Microsoft Graph client in Node.js environment
require('isomorphic-fetch');

const { Client } = require('@microsoft/microsoft-graph-client');

// --- searchMessageByMetadata: Primary search logic using subject and recipients ---
async function searchMessageByMetadata(client, subjectFromRequest, recipientsFromRequest, context) {
    // Helper to encode strings for OData filters and trim them.
    const encodeOData = (str) => str ? str.replace(/'/g, "''").trim() : "";

    const initialODataFilterParts = [];

    // Add subject to filter if provided and not an empty string after trimming.
    const encodedSubject = encodeOData(subjectFromRequest);
    if (encodedSubject) {
        initialODataFilterParts.push(`subject eq '${encodedSubject}'`);
    }

    // Process recipients string into a list of lowercase, trimmed, non-empty email addresses.
    const fullRecipientList = recipientsFromRequest 
        ? recipientsFromRequest.split(';').map(r => r.trim().toLowerCase()).filter(r => r) 
        : [];

    // If recipients are provided, use the first valid one in the initial OData filter.
    // This helps narrow down results if the subject is too generic or missing.
    if (fullRecipientList.length > 0) {
        const firstRecipientEncoded = encodeOData(fullRecipientList[0]); // Already lowercased by previous map
        if (firstRecipientEncoded) { // Ensure it's not empty after encoding (e.g. if original was just "'")
            initialODataFilterParts.push(`toRecipients/any(r: r/emailAddress/address eq '${firstRecipientEncoded}')`);
        }
    }
    
    // If no subject AND no recipients were provided to filter by, the search is too broad.
    // The function will throw an error to prevent an overly broad or potentially erroneous query.
    if (initialODataFilterParts.length === 0) {
        if (context && context.log) {
            context.log.warn("searchMessageByMetadata: Called without subject or any recipients for initial filter. Search would be too broad and is aborted.");
        }
        throw new Error("Search requires at least a subject or recipients to be specified for filtering.");
    }

    const initialODataFilter = initialODataFilterParts.join(' and ');
    if (context && context.log) context.log.info(`searchMessageByMetadata: Constructing OData filter for Graph API: ${initialODataFilter}`);

    try {
        // Fetch top N messages (e.g., 5) matching the simplified initial filter, ordered by most recent.
        // 'toRecipients' and 'subject' are selected for secondary validation post-fetch.
        const response = await client.api('/me/messages')
            .filter(initialODataFilter)
            .orderby('receivedDateTime desc') // Get the latest messages first
            .top(5) // Fetch a small batch for secondary filtering, reducing API load.
            .select('id,receivedDateTime,subject,toRecipients')
            .get();

        if (response.value && response.value.length > 0) {
            if (context && context.log) context.log.info(`searchMessageByMetadata: Initial query returned ${response.value.length} message(s). Performing detailed client-side validation...`);
            
            // Iterate through the fetched messages for detailed validation.
            for (const message of response.value) {
                // Secondary Validation:
                // 1. Validate Subject:
                //    If a subject was provided in the request, ensure the message's subject matches exactly (case-insensitive).
                let subjectMatch = true; // Assume match if no subject was in the original request to filter by.
                if (encodedSubject) { // Only check if a subject filter was applied.
                    const messageSubjectNormalized = message.subject ? message.subject.trim().toLowerCase() : "";
                    // Compare with the normalized subject from the request.
                    subjectMatch = (messageSubjectNormalized === encodedSubject.toLowerCase()); 
                    if (!subjectMatch && context && context.log) {
                        context.log.info(`searchMessageByMetadata: Message ${message.id} (Received: ${message.receivedDateTime}): Subject mismatch post-query. Expected (normalized): "${encodedSubject.toLowerCase()}", Actual (normalized): "${messageSubjectNormalized}".`);
                    }
                }
                if (!subjectMatch) continue; // If subject doesn't align, skip to the next message.

                // 2. Validate All Recipients:
                //    If recipients were specified in the request, ensure ALL of them are present in this message's 'toRecipients'.
                let recipientsMatch = true; // Assume match if no recipients were in the original request.
                if (fullRecipientList.length > 0) { // Only validate if recipients were expected.
                    recipientsMatch = false; // Reset to false, needs to be proven true.
                    if (message.toRecipients && message.toRecipients.length > 0) {
                        // Create a Set of recipients from the current message for efficient lookup.
                        const messageRecipientsSet = new Set(
                            message.toRecipients.map(r => r.emailAddress && r.emailAddress.address ? r.emailAddress.address.toLowerCase() : null).filter(Boolean)
                        );
                        
                        let allExpectedRecipientsFound = true;
                        for (const expectedRecipient of fullRecipientList) { // fullRecipientList is already lowercased.
                            if (!messageRecipientsSet.has(expectedRecipient)) {
                                allExpectedRecipientsFound = false; // An expected recipient is missing.
                                if (context && context.log) {
                                    context.log.info(`searchMessageByMetadata: Message ${message.id} (Received: ${message.receivedDateTime}): Recipient mismatch. Expected recipient "${expectedRecipient}" not found in message recipients: [${Array.from(messageRecipientsSet).join(', ')}].`);
                                }
                                break; // No need to check further recipients for this message.
                            }
                        }
                        if (allExpectedRecipientsFound) recipientsMatch = true; // All expected recipients were found.
                    } else if (context && context.log) { // Message has no recipients, but we expected some.
                        context.log.info(`searchMessageByMetadata: Message ${message.id} (Received: ${message.receivedDateTime}): Recipient mismatch. Expected ${fullRecipientList.length} recipients, but message has none.`);
                    }
                }
                if (!recipientsMatch) continue; // If recipients don't align, skip to the next message.


                // If all validations pass, this is the latest message matching all criteria.
                if (context && context.log) context.log.info(`searchMessageByMetadata: SUCCESS - Message ${message.id} (Received: "${message.receivedDateTime}", Subject: "${message.subject}") passed all client-side validations.`);
                return message.id; // Return the ID of the validated message.
            }
            // If the loop completes, no message passed all detailed validations.
            if (context && context.log) context.log.info("searchMessageByMetadata: No messages passed detailed client-side validation after initial query.");
        } else if (context && context.log) { // Initial query returned no messages.
            context.log.info(`searchMessageByMetadata: Initial OData query returned no messages. Filter used: "${initialODataFilter}"`);
        }
        return null; // No message found matching all criteria.
    } catch (err) {
        if (context && context.log) context.log.error(`searchMessageByMetadata: Error during Graph API call (Filter was: "${initialODataFilter}"): ${err.message}`);
        throw new Error(`Error in searchMessageByMetadata: ${err.message}`); // Propagate the error.
    }
}

// --- Main Azure Function Handler ---
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

        // Destructure payload from request body. ewsItemId is no longer expected.
        // receivedTime is still received but not directly used in searchMessageByMetadata's filter.
        const {
            subject,
            recipients, // Expected to be a string (potentially empty) from command.html
            // receivedTime, // Not directly passed to the new searchMessageByMetadata filter logic
            // userEmail,    // Not used in core logic here
        } = req.body || {};
        
        let messageIdToProcess = null;

        // Directly use metadata search as the only method.
        // `useMetadataSearch` flag from payload is implicitly true.
        context.log.info(`Attempting metadata search. Criteria: Subject="${subject}", Recipients="${recipients}"`);
        try {
            // Call searchMessageByMetadata. It will throw if subject and recipients are both missing/empty.
            // The `recipients` variable passed here is the string from the request body.
            messageIdToProcess = await searchMessageByMetadata(client, subject, recipients, context);
            
            if (messageIdToProcess) {
                context.log(`Metadata search successful. Found message with Graph REST ID: "${messageIdToProcess}".`);
            } else {
                // If searchMessageByMetadata returns null, it means no message matched all criteria.
                const recipientsForLog = typeof recipients === 'string' ? recipients : JSON.stringify(recipients);
                context.log.error(`Metadata search did not find a matching message. Criteria: Subject="${subject}", Recipients="${recipientsForLog}"`);
                context.res = { status: 404, body: { success: false, error: `No message found via metadata search matching criteria (Subject: "${subject}", Recipients: "${recipientsForLog}")` } };
                return;
            }
        } catch (searchError) {
            // Handle errors from searchMessageByMetadata (e.g., missing criteria, Graph API call failure).
            context.log.error(`Error during metadata search: ${searchError.message}`);
            context.res = { status: 500, body: { success: false, error: `Error searching for message via metadata: ${searchError.message}` } };
            return;
        }
        
        // At this point, messageIdToProcess should be a valid Graph REST ID.
        context.log(`Proceeding to process message with Graph REST ID: "${messageIdToProcess}"`);
        
        // --- Message Processing (Fetch, Draft, Attachments, Send, Move) ---
        // This section remains largely the same, using messageIdToProcess.
        try {
            const message = await client.api(`/me/messages/${messageIdToProcess}`)
                .select('subject,body,toRecipients,ccRecipients,bccRecipients,from,hasAttachments,importance,isRead')
                .get();
            context.log(`Successfully retrieved original message: "${message.subject}" (ID: ${message.id})`);

            let attachments = [];
            if (message.hasAttachments) {
                context.log("Original message has attachments. Fetching attachment details...");
                const attachmentsResponse = await client.api(`/me/messages/${messageIdToProcess}/attachments`).get();
                attachments = attachmentsResponse.value || []; // Ensure attachments is an array.
                context.log(`Found ${attachments.length} attachments in the original message.`);
            }

            context.log("Creating new message draft for forwarding...");
            const newMessage = {
                subject: `${message.subject}`, // Preserve original subject.
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
                        // Populate specific properties based on attachment type.
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
                             continue; // Skip this attachment.
                        } else if(!attachmentData.contentBytes && !attachmentData.item && !attachmentData.sourceUrl && 
                                  (attachment["@odata.type"] === "#microsoft.graph.fileAttachment" || 
                                   attachment["@odata.type"] === "#microsoft.graph.itemAttachment" || 
                                   attachment["@odata.type"] === "#microsoft.graph.referenceAttachment") ) { 
                            // Known type but essential data missing.
                            context.log.warn(`Attachment "${attachment.name}" (Type: ${attachment["@odata.type"]}) is a known type but missing required data (e.g., contentBytes, item, sourceUrl). Skipping.`);
                            continue; // Skip this attachment.
                        }

                        await client.api(`/me/messages/${draftMessage.id}/attachments`).post(attachmentData);
                        context.log(`Successfully added attachment "${attachment.name}" to draft ${draftMessage.id}.`);
                    } catch (attachError) {
                        const errBody = attachError.body ? JSON.stringify(attachError.body) : 'N/A';
                        context.log.error(`Error adding attachment "${attachment.name}" to draft ${draftMessage.id}: ${attachError.message}. Error Body: ${errBody}`);
                        // Continue with other attachments even if one fails.
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
            // Handle errors related to fetching/processing the message *after* its ID is known.
            context.log.error(`Error during message processing/forwarding for Graph REST ID "${messageIdToProcess}": ${processingError.message}`);
            let errMsg = `Error processing message (ID: "${messageIdToProcess}"): ${processingError.message}`;
            if (processingError.statusCode && processingError.code) { // Check if it's a Graph API error object.
                errMsg = `Graph API Error (${processingError.code}) for message ID "${messageIdToProcess}": ${processingError.message}`;
            }
            context.res = { status: (processingError.statusCode === 404 ? 404 : 500), body: { success: false, error: errMsg, messageIdUsed: messageIdToProcess } };
        }
    } catch (error) { // Outermost catch block for any unhandled errors.
        context.log.error(`Unhandled error in Azure Function: ${error.message}`);
        context.res = { status: 500, body: { success: false, error: `Critical error in email forwarding process: ${error.message}` } };
    }
};

// --- Helper Functions ---
function getAuthenticatedClient(accessToken) {
    const client = Client.init({ authProvider: (done) => done(null, accessToken) });
    return client;
}

// convertExchangeId and validateMessageId functions are no longer needed as EWS ID is not used.
