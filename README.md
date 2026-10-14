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

## Next steps (not yet built)

- Pull the signature template + per-user data from SharePoint (as the WPF app does) instead of the hardcoded prototype signature.
- Decide hosting for real deployment (GitHub Pages / static host) and central deployment via the M365 admin center.
