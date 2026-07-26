let mailboxItem = null;
let completed = false; 

Office.initialize = function (reason)
{
    console.log("*******************************************************");
    console.log("Office.initialize");
    mailboxItem = Office.context.mailbox.item;
}

function getDialogUrl()
{
    return "https://danilkorotenko.github.io/OutlookAddin_dialogUI/dialog.html";
}

function completeBlocked()
{
    if (completed)
    {
        return;
    }
    completed = true;
    event.completed({ allowEvent: false });
}


function validateMessage(event)
{
    console.log("Start validation stream");

    Office.context.ui.displayDialogAsync(
        getDialogUrl(),
        { height: 30, width: 30, displayInIframe: true },
        function (asyncResult)
        {
            if (asyncResult.status === Office.AsyncResultStatus.Failed)
            {
                console.error("Failed to open dialog: " + asyncResult.error.message);
                completeBlocked();
                return;
            }

            const dialog = asyncResult.value;

            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function ()
            {
                dialog.close();
                completeBlocked();
            });

            dialog.addEventHandler(Office.EventType.DialogEventReceived, function ()
            {
                completeBlocked();
            });
        });
}
