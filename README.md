# Outlook Signature Add-in (prototype)

An Outlook **web add-in** that automatically inserts a company email signature when you
compose a **new message**. Uses event-based activation (`OnNewMessageCompose`) and the
`setSignatureAsync` API — the only Microsoft-supported way to set signatures in the
**new Outlook for Windows** (also works in classic Outlook 2201+, Outlook on the web, and Mac).

This is a separate project from the WPF **Signature Manager** (which writes the local
`AppData\Roaming\Microsoft\Signatures` folder and therefore only covers *classic* Outlook).

## Prerequisites

- Node.js (installed)
- Outlook (new or classic) signed in to a Microsoft 365 mailbox

## Quick start (local dev)

```bash
npm install
npm run certs      # one-time: install the trusted localhost HTTPS cert (asks for admin)
npm run icons      # generate the PNG icons
npm run validate   # sanity-check the manifest
npm run dev        # serve the add-in at https://localhost:3000  (keep running)
```

Then, in a second terminal, sideload into desktop Outlook and restart it:

```bash
npm run sideload   # registers manifest.xml for developer sideloading
```

Start a **new message** in Outlook — the signature is inserted automatically.

To remove: `npm run unsideload`.

## Files

| File | Purpose |
|------|---------|
| `manifest.xml` | Add-in manifest; event-based activation + runtime URLs |
| `src/commands/commands.html` | Hosts the Office.js runtime (no UI) |
| `src/commands/commands.js` | The `onNewMessageComposeHandler` that builds + sets the signature |
| `dev-server.js` | HTTPS static server on :3000 |
| `assets/make-icons.js` | Generates solid-color PNG icons |
| `sideload.ps1` | Registers/removes the manifest for desktop sideloading |

## Hosting (GitHub Pages)

The add-in is hosted at **https://marcbuetti.github.io/signature-addin/** (GitHub Pages, master branch).
The manifest points at these URLs. To update the add-in, push to `master`; Pages redeploys automatically.

- Runtime: https://marcbuetti.github.io/signature-addin/commands/commands.html
- Manifest: https://marcbuetti.github.io/signature-addin/manifest.xml

## Deployment to new Outlook (M365 admin center)

New Outlook and Outlook on the web no longer offer end-user "add from file" sideloading, so the add-in
is deployed centrally:

1. https://admin.microsoft.com → **Settings → Integrated apps → Upload custom apps**
2. App type **Office Add-in**, provide the manifest URL above (or upload `manifest.xml`)
3. Assign to a test user / group, accept permissions, deploy
4. It appears in new Outlook within minutes–hours; start a new message to trigger the signature

(Local dev with `npm run dev` + registry sideload still works for **classic** Outlook testing.)

## Next steps (not yet built)

- Pull the signature template + per-user data from SharePoint (as the WPF app does) instead of the hardcoded prototype signature.
