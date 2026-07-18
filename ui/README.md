# KAAVIS Voice Shell (Phase 0)

A single self-contained HTML file (`index.html`) implementing the voice-first control surface described in [`docs/PRD.md`](../docs/PRD.md) — Section 8 (UX Principles) and Section 10 (Permission Model).

## What this is

- A working voice loop: real speech-to-text (Web Speech API) and text-to-speech (`speechSynthesis`), no server required.
- The HUD orb reflects live state — idle, listening (reacts to your mic's volume), speaking.
- An **Integrations** panel listing every source from PRD Section 6 (Email, Word, Google Docs, Obsidian, Apple Notes, Safari), each honestly marked as not connected, with what it needs to go live.
- An **Action Log** — the accountability record described in PRD Section 10, currently logging voice input, KAAVIS's replies, and connect requests.
- An **output voice picker**, populated from your system's installed voices, plus a placeholder for a custom TTS voice (provider + voice ID) once one is wired up.

## What this isn't (yet)

No integration in the left panel is actually connected — there's no backend, no OAuth, no file access. Clicking "Connect" logs the request and tells you what's needed; it doesn't perform the connection. That work follows the build sequence in PRD Section 12, starting with local file + Obsidian vault access.

## Running it

Open `index.html` directly in Chrome or Edge (voice input needs the Web Speech API, which Firefox/Safari don't support). No build step, no dependencies — fonts are embedded, everything runs client-side.

## Adding a voice

Open the "Custom voice" panel at the bottom right. Pick a provider (ElevenLabs, Azure, Play.ht, or Other) and paste a voice ID — it's saved locally in your browser only. Nothing is sent anywhere until that provider's API is actually wired up.
