# KAAVIS Voice Shell (Phase 0)

A single self-contained HTML file (`index.html`) implementing the voice-first control surface described in [`docs/PRD.md`](../docs/PRD.md) — Section 8 (UX Principles) and Section 10 (Permission Model).

## What this is

- A working voice loop: real speech-to-text (Web Speech API or Deepgram) and text-to-speech (`speechSynthesis`).
- The HUD orb reflects live state — idle, listening (reacts to your mic's volume), speaking.
- A typed-command box below the transcript, for when a mic isn't handy — goes through the exact same command handling as voice.
- **Two things that actually execute**, no credentials needed:
  - **Web Search** — "search for `<query>`" opens real results in a new browser tab.
  - **Open Apps** — "open `<app or site>`" launches a local app (via a small companion server, see below) or opens a website, depending on what you say.
- An **Integrations** panel: Web Search and Open Apps show as connected; Email, Word, Google Docs, Apple Notes, and Safari are honestly marked not connected, with what each needs to go live (real OAuth credentials or platform bridges I can't fabricate on your behalf).
- An **Action Log** — logs voice input, KAAVIS's replies, and every action taken.
- An **output voice picker**, populated from your system's installed voices, plus a custom-voice placeholder (provider + voice ID) for a future TTS provider.

## What this isn't (yet)

Email, Word, Google Docs, Apple Notes, and Safari still need real setup — Gmail/Google Docs need an OAuth app you create yourself, Apple Notes/Safari need the iPad Shortcuts relay described in `docs/PRD.md` Section 6. I can't fabricate credentials on your behalf; clicking "Connect" on those rows tells you what's needed rather than faking a connection.

## Running it

**Open apps requires the local server** (it calls a companion API that isn't available if you just double-click the file). Start it with:

```
powershell -ExecutionPolicy Bypass -File .claude/serve-ui.ps1
```

then open **http://localhost:5550/** in Chrome or Edge. Web search, voice, and everything else still work if you just double-click `ui/index.html` directly — only "open an app" needs the server running.

**A note on "Open Apps"**: the HTTP API and its success/failure reporting are verified correct — it genuinely calls `Start-Process` and returns real results (confirmed a bad app name correctly returns an honest error). In my own test environment, apps launched through the server sometimes didn't stay open, while the identical command run directly in an interactive terminal worked reliably — this looks like an artifact of how my automated tooling supervises background processes, not a bug in the script itself, but I couldn't fully root-cause it. Please verify on your machine (run the server from your own terminal, not through any automation, and try "open notepad") before relying on it.

## Wake word

Push-to-talk (clicking the mic) always acts on whatever it hears. **Always listening** mode is different: with a wake word set (the "Wake word" field next to the toggle, default `kaavis`), it ignores ambient speech and only responds to phrases containing that word — say "hey kaavis, open notepad" and it strips the wake word and runs "open notepad". Clear the field to make Always listening act on everything again (not recommended unless you're in a quiet room).

Commands also tolerate a bit of natural phrasing now — leading "hey", "please", "can you", etc. are stripped before matching, so "can you search for pizza please" works the same as "search for pizza".

## Adding a voice

Open the "Custom voice" panel at the bottom right. Pick a provider (ElevenLabs, Azure, Play.ht, or Other) and paste a voice ID — it's saved locally in your browser only. Nothing is sent anywhere until that provider's API is actually wired up.

## Using Deepgram for voice input

The default speech-to-text engine is the browser's built-in `SpeechRecognition` (Chrome/Edge only, no custom vocabulary). Open the "Speech input" panel next to the mic button, switch the engine to **Deepgram**, and paste an API key from [console.deepgram.com](https://console.deepgram.com) — this gives real-time streaming transcription with better accuracy across more browsers.

The key is saved to this browser's local storage only and sent directly from your browser to Deepgram over a WebSocket. It is **never** written to this repo or sent anywhere else — `kaavis` is a public repo, so no key should ever be hardcoded into `index.html`. Because KAAVIS is a single-user personal tool, storing your own key in your own browser is an accepted tradeoff; don't share that browser profile or its local storage.

This path was reviewed carefully but not run end-to-end against a live Deepgram account — the sandboxed environment used to build it can't grant microphone access. Test it yourself with a real key before relying on it.
