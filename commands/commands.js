/*
 * Event-based handler: on new message compose, insert the default signature.
 * Shows a progress indicator while it works, so the brief runtime cold-start has visible feedback.
 */

/* global Office, SignatureLib */

Office.onReady(function () {});

function onNewMessageComposeHandler(event) {
  var item = Office.context.mailbox.item;

  // Visible "loading" feedback at the top of the compose window (spinner).
  try {
    item.notificationMessages.addAsync("sigLoading", {
      type: "progressIndicator",
      message: "Signatur wird eingefügt…"
    });
  } catch (e) { /* notifications unavailable — continue silently */ }

  var html = window.SignatureLib.build(window.SignatureLib.defaultId, Office.context.mailbox.userProfile);

  item.body.setSignatureAsync(
    html,
    { coercionType: Office.CoercionType.Html, asyncContext: event },
    function (res) {
      try { item.notificationMessages.removeAsync("sigLoading"); } catch (e) {}
      if (res.status === Office.AsyncResultStatus.Failed) {
        try {
          item.notificationMessages.addAsync("sigError", {
            type: "errorMessage",
            message: "Signatur konnte nicht eingefügt werden."
          });
        } catch (e) {}
      }
      res.asyncContext.completed();
    }
  );
}

if (Office.actions && typeof Office.actions.associate === "function") {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
