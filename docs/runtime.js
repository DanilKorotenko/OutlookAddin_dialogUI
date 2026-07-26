let mailboxItem = null;

Office.initialize = function (reason)
{
    console.log("*******************************************************");
    console.log("Office.initialize");
    mailboxItem = Office.context.mailbox.item;
}

function showBlockedDialog(onClosed)
{
    Office.context.ui.displayDialogAsync(
        "https://danilkorotenko.github.io/OutlookAddin_dialogUI/dialog.html",
        { height: 30, width: 30, promptBeforeOpen: false, displayInIframe: true },
        function (asyncResult)
        {
            if (asyncResult.status === Office.AsyncResultStatus.Failed)
            {
                console.error("Failed to open dialog: " + asyncResult.error.message);
                if (onClosed)
                {
                    onClosed();
                }
                return;
            }

            const dialog = asyncResult.value;
            let closed = false;

            function finish()
            {
                if (closed)
                {
                    return;
                }
                closed = true;
                if (onClosed)
                {
                    onClosed();
                }
            }

            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function ()
            {
                console.log("Dialog event DialogMessageReceived");
                dialog.close();
                finish();
            });

            dialog.addEventHandler(Office.EventType.DialogEventReceived, function ()
            {
                console.log("Dialog event DialogEventReceived");
                finish();
            });
        });
}

function validateMessage(event)
{
    console.log("Start validation stream");

    showBlockedDialog(function ()
    {
        event.completed({ allowEvent: false });
    });
}
