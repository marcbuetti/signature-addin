/*
 * Event-based activation handler.
 * Fires on "new message compose" and sets the signature via setSignatureAsync.
 *
 * In a real build the signature template + user data would come from the
 * SharePoint profile (as the WPF Signature Manager does today). For this
 * prototype we build it from the signed-in user's Office profile.
 */

/* global Office */

Office.onReady(() => {
  // Nothing to do on load; the handler is associated below.
});

/**
 * Builds an email-safe (inline-styled) HTML signature from the user profile.
 */
function buildSignatureHtml(profile) {
  const name = (profile && profile.displayName) || "";
  const email = (profile && profile.emailAddress) || "";
  const company = "Vegilife AG";
  const brand = "#3BA776";

  return [
    '<div style="font-family:Segoe UI, Arial, sans-serif; font-size:11pt; color:#222; line-height:1.35;">',
    '  <div style="margin-top:10px; border-top:3px solid ', brand, '; width:260px;"></div>',
    '  <div style="margin-top:8px; font-weight:600; font-size:12pt; color:#111;">', escapeHtml(name), '</div>',
    '  <div style="color:', brand, '; font-weight:600;">', escapeHtml(company), '</div>',
    email
      ? '  <div style="margin-top:4px;">E-Mail: <a href="mailto:' + escapeAttr(email) + '" style="color:' + brand + '; text-decoration:none;">' + escapeHtml(email) + '</a></div>'
      : '',
    '  <div style="margin-top:2px;">Web: <a href="https://www.vegilife.ch" style="color:', brand, '; text-decoration:none;">www.vegilife.ch</a></div>',
    '</div>',
  ].join("");
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}
function escapeAttr(s) {
  return escapeHtml(s);
}

/**
 * The function referenced by <LaunchEvent FunctionName="onNewMessageComposeHandler"> in the manifest.
 */
function onNewMessageComposeHandler(event) {
  const profile = Office.context.mailbox.userProfile;
  const signatureHtml = buildSignatureHtml(profile);

  Office.context.mailbox.item.body.setSignatureAsync(
    signatureHtml,
    { coercionType: Office.CoercionType.Html, asyncContext: event },
    (asyncResult) => {
      // Always complete the event so Outlook stops showing the "running" state.
      const ev = asyncResult.asyncContext;
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("setSignatureAsync failed:", asyncResult.error && asyncResult.error.message);
      }
      ev.completed();
    }
  );
}

// Associate the handler for event-based activation (required so Outlook can find it by name).
if (Office.actions && typeof Office.actions.associate === "function") {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
