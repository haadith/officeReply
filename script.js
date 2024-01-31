Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Office is ready
        document.getElementById("replyWithAttachment").onclick = () => replyWithAttachment();
    }
});

async function replyWithAttachment() {
    try {
        await Office.context.mailbox.item.body.getAsync("html", { asyncContext: "This is passed to the callback" }, function (result) {
            let emailBody = result.value; // Original email body in HTML
            let replyBody = "<p>Your custom reply message here...</p>" + emailBody; // Prepend your custom message

            Office.context.mailbox.item.displayReplyAllForm({
                'htmlBody': replyBody,
                'attachments': getAttachmentsFromItem(Office.context.mailbox.item)
            });
        });
    } catch (error) {
        console.error(error);
    }
}

function getAttachmentsFromItem(item) {
    let attachments = [];
    if (item.attachments && item.attachments.length > 0) {
        item.attachments.forEach((attachment) => {
            attachments.push({
                "type": "file",
                "name": attachment.name,
                "url": attachment.url,
                "isInline": attachment.isInline
            });
        });
    }
    return attachments;
}
