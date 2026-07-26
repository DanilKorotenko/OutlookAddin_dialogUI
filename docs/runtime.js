let mailboxItem = null;

Office.initialize = function (reason)
{
    console.log("*******************************************************");
    console.log("Office.initialize");
    mailboxItem = Office.context.mailbox.item;
}

function validateMessage(event)
{
    console.log("Start validation stream");

    Office.context.ui.displayDialogAsync(
        "https://danilkorotenko.github.io/OutlookAddin_dialogUI/dialog.html",
        // { height: 30, width: 30, displayInIframe: true },
        { height: 30, width: 30, promptBeforeOpen: false, displayInIframe: true, },
        function (asyncResult)
        {
            if (asyncResult.status === Office.AsyncResultStatus.Failed)
            {
                console.error("Failed to open dialog: " + asyncResult.error.message);
                event.completed({ allowEvent: false });
                return;
            }

            const dialog = asyncResult.value;

            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function ()
            {
                dialog.close();
                event.completed({ allowEvent: false });
            });

            dialog.addEventHandler(Office.EventType.DialogEventReceived, function ()
            {
                event.completed({ allowEvent: false });
            });
        });
}
