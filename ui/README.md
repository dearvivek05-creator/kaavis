# KAAVIS Voice Shell (Phase 0)

A single self-contained HTML file (`index.html`) implementing the voice-first control surface described in [`docs/PRD.md`](../docs/PRD.md) — Section 8 (UX Principles) and Section 10 (Permission Model).

## What this is

- A working voice loop: real speech-to-text (Web Speech API) and text-to-speech (`speechSynthesis`), no server required.
- A typed-command fallback (below the transcript) that goes through the exact same command handling as voice — useful when a mic isn't available.
- The HUD orb reflects live state — idle, listening (reacts to your mic's volume), speaking.
- An **Integrations** panel listing every source from PRD Section 6 (Email, Word, Google Docs, Obsidian, Apple Notes, Safari), each honestly marked as not connected, with what it needs to go live.
- **Obsidian Vault is actually wired up** — via the browser's File System Access API, no server or OAuth needed. Click Connect on that row, pick your vault folder, and KAAVIS can:
  - `list my notes` / `what are my notes`
  - `read my <name> note` / `what does my <name> note say`
  - `create a note called <name> saying <content>` (add ` in <folder>` to nest it, e.g. "create a note called KAAVIS in Projects saying ...")
  - `add <content> to my <name> note` (appends)
  - `search my vault for <query>` / `what did I write about <query>`
  - The folder handle is remembered (IndexedDB) so you don't have to re-pick it every session, though Chrome will ask you to re-grant permission on reload.
- An **Action Log** — the accountability record described in PRD Section 10, currently logging voice input, KAAVIS's replies, connect requests, and vault reads/writes.
- An **output voice picker**, populated from your system's installed voices, plus a placeholder for a custom TTS voice (provider + voice ID) once one is wired up.

## What this isn't (yet)

Every integration except Obsidian Vault is still a placeholder — no backend, no OAuth, no file access. Clicking "Connect" on those logs the request and tells you what's needed; it doesn't perform the connection. That work follows the build sequence in PRD Section 12.

The vault integration is real but only mechanical: it can create, read, append, search, and list notes exactly as told. It can't summarize a conversation or reason about content — that needs an LLM backend, which isn't connected. Ask it to "summarize" into a note and it'll say so rather than writing something fake.

**Not yet verified against a real vault** — the read/write logic was reviewed carefully but the sandboxed test environment used to build this can't drive the native folder-picker dialog, so please test the Connect flow yourself with an actual Obsidian vault (or any folder of `.md` files) before relying on it.

## Running it

Open `index.html` directly in Chrome or Edge (voice input needs the Web Speech API, which Firefox/Safari don't support). No build step, no dependencies — fonts are embedded, everything runs client-side.

## Adding a voice

Open the "Custom voice" panel at the bottom right. Pick a provider (ElevenLabs, Azure, Play.ht, or Other) and paste a voice ID — it's saved locally in your browser only. Nothing is sent anywhere until that provider's API is actually wired up.

## Using Deepgram for voice input

The default speech-to-text engine is the browser's built-in `SpeechRecognition` (Chrome/Edge only, no custom vocabulary). Open the "Speech input" panel next to the mic button, switch the engine to **Deepgram**, and paste an API key from [console.deepgram.com](https://console.deepgram.com) — this gives real-time streaming transcription with better accuracy across more browsers.

The key is saved to this browser's local storage only and sent directly from your browser to Deepgram over a WebSocket. It is **never** written to this repo or sent anywhere else — `kaavis` is a public repo, so no key should ever be hardcoded into `index.html`. Because KAAVIS is a single-user personal tool, storing your own key in your own browser is an accepted tradeoff; don't share that browser profile or its local storage.

This path was reviewed carefully but not run end-to-end against a live Deepgram account — the sandboxed environment used to build it can't grant microphone access. Test it yourself with a real key before relying on it.
