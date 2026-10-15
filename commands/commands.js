/*
 * Event handlers:
 *  - OnNewMessageCompose: insert the default signature (with a progress indicator).
 *  - OnMessageSend:       re-apply the authoritative signature right before sending, so any manual
 *                         edits to the signature are overwritten (tamper-proof at send time).
 * The chosen template id is remembered per message via item custom properties, so switching to
 * "Team"/"Kurz" in the task pane is preserved on send.
 */

/* global Office */

Office.onReady(function () {});

var DEFAULT_ID = (window.SignatureLib && window.SignatureLib.defaultId) || "standard";

function saveTemplate(item, id, done) {
  try {
    item.loadCustomPropertiesAsync(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
        r.value.set("sigTemplate", id);
        r.value.saveAsync(function () { if (done) done(); });
      } else if (done) { done(); }
    });
  } catch (e) { if (done) done(); }
}

function readTemplate(item, cb) {
  try {
    item.loadCustomPropertiesAsync(function (r) {
      var id = DEFAULT_ID;
      if (r.status === Office.AsyncResultStatus.Succeeded && r.value && r.value.get("sigTemplate")) {
        id = r.value.get("sigTemplate");
      }
      cb(id);
    });
  } catch (e) { cb(DEFAULT_ID); }
}

function applySignature(item, id, done) {
  var html = window.SignatureLib.build(id, Office.context.mailbox.userProfile);
  item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, function (res) {
    if (done) done(res);
  });
}

// --- Compose: insert default signature + loading indicator ---
function onNewMessageComposeHandler(event) {
  var item = Office.context.mailbox.item;
  try {
    item.notificationMessages.addAsync("sigLoading", {
      type: "progressIndicator",
      message: "Signatur wird eingefügt…"
    });
  } catch (e) {}

  applySignature(item, DEFAULT_ID, function (res) {
    try { item.notificationMessages.removeAsync("sigLoading"); } catch (e) {}
    saveTemplate(item, DEFAULT_ID, function () { event.completed(); });
  });
}

// --- Send: re-apply the authoritative signature, then allow the send ---
function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  readTemplate(item, function (id) {
    applySignature(item, id, function () {
      event.completed({ allowEvent: true });
    });
  });
}

// Let the task pane record which template the user picked (so send re-applies the same one).
window.rememberTemplate = function (id) {
  saveTemplate(Office.context.mailbox.item, id);
};

if (Office.actions && typeof Office.actions.associate === "function") {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
