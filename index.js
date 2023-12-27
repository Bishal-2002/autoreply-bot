require('dotenv').config()

// This package is imported from the googleapis module and provides the necessary functionality to interact with various Google APIs, including the Gmail API.
const {google} = require('googleapis')

const oAuth2Client = new google.auth.OAuth2(
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET,
    process.env.REDIRECT_URI
)
oAuth2Client.setCredentials({ refresh_token: process.env.REFRESH_TOKEN })

// This method checks and sends automated reply to all unread messages which doesn't have any prior reply.
const checkEmailsAndSendReply = async () => {
    try {
        const gmail = google.gmail({ version: "v1", auth: oAuth2Client })

        // Getting the list of unread mails
        const res = await gmail.users.messages.list({
            userId: "me",
            q: "is:unread",
        })
        const messages = res.data.messages

        // If there are messages to work with
        if (messages && messages.length > 0) {

            // Getting the message details
            for (const message of messages) {

                const email = await gmail.users.messages.get({
                    userId: "me",
                    id: message.id
                })

                const from = email.data.payload.headers.find(
                    (header) => header.name === "From"
                )
                const toHeader = email.data.payload.headers.find(
                    (header) => header.name === "To"
                )
                const Subject = email.data.payload.headers.find(
                    (header) => header.name === "Subject"
                )

                // Finding the unique message id
                const messageID = email.data.payload.headers.find(
                    (header) => header.name === "Message-ID"
                ).value

                // Email ID of sender
                const fromEmail = from.value
                // Email ID of receiver
                const toEmail = toHeader.value
                // Subject of mail
                const subject = Subject.value

                // Checking if the mail has prior replies
                const threadID = message.threadId
                const thread = await gmail.users.threads.get({
                    userId: "me",
                    id: threadID,
                })

                // Getting replies
                const replies = thread.data.messages.slice(1)
                if (replies.length === 0) {

                    // Replying to the email
                    await gmail.users.messages.send({
                        userId: "me",
                        requestBody: {
                            id: messageID,
                            threadId: threadID,
                            raw: await createEncodedReply(toEmail, fromEmail, `Re: ${subject}`, messageID)
                        }
                    })

                    // Adding label
                    const labelName = "Auto Reply - Vacation"
                    await gmail.users.messages.modify({
                        userId: "me",
                        id: message.id,
                        requestBody: {
                            addLabelIds: [await createLabel(labelName)],
                            removeLabelIds: ['INBOX']
                        }
                    })

                    console.log(`Sent reply to ${fromEmail}`)
                }                
            }
        }
    }
    catch (err) { console.log(err) }
}

// It converts String to Base64Encoded format
const createEncodedReply = (from, to, subject, messageID) => {
    const emailContent = `From: ${from}\r\nTo: ${to}\r\nSubject: ${subject}\r\nIn-Reply-To: <${messageID}>\r\nReferences: <${messageID}>\r\nMessage-ID: ${messageID}\r\n\r\nHi,\n\nThanks for reaching out. I am currently on a vacation. I will revert back to you once I return back home.\n\nThanks & Regards,\nBishal`

    const base64EncodedEmail = Buffer.from(emailContent).toString("base64")
  
    return base64EncodedEmail;
}

// This method creates the label
const createLabel = async (labelName) => {
    const gmail = google.gmail({ version: "v1", auth: oAuth2Client });

    // Checking if label already exists!
    const res = await gmail.users.labels.list({ userId: "me" });
    const labels = res.data.labels;
    const existingLabel = labels.find((label) => label.name === labelName);
    if (existingLabel)      return existingLabel.id;
  
    // Creating the label if it doesn't exist.
    const newLabel = await gmail.users.labels.create({
        userId: "me",
        requestBody: {
            name: labelName,
            labelListVisibility: "labelShow",
            messageListVisibility: "show"
        }
    })
  
    return newLabel.data.id;
}

// Helper function to sleep
const sleep = (time) => {
    return new Promise(resolve => setTimeout(resolve, time));
}

async function main() {
    try {
        while (true) {
            await checkEmailsAndSendReply()
            console.log('Done!')
            const interval = Math.floor(Math.random() * (120 - 45 + 1) + 45)
            console.log(`Waiting for ${interval} seconds...`)
            await sleep(interval * 1000) // convert s to ms
        }
    }
    catch (err) { console.log(err) }
}

// Checking if this file is the entry point
if (require.main === module) {
    main();
}

/*
    Areas of improvement --
    1. Error handling: The code currently logs any errors that occur during the execution but does not handle them in a more robust manner.
    2. Code efficiency: The code could be optimized to handle larger volumes of emails more efficiently.
    3. User Login Option: Currently, this code uses a single gmail account which is already hard coded, to check emails. A Google Login can be implemented which allows the user to login into his/her gmail account and use this bot for their account.
    4. User-specific configuration: Making the code more flexible by allowing users to provide their own configuration options, such as email filters or customized reply messages. 
    5. Time Monitoring: The code currently use random interval function to generate seconds and in this code can be improved by adding cron jobs package to schedule email tasks 

    These are some areas where the code can be improved, but overall, it provides implementation of auto-reply functionality using the Gmail API.
*/