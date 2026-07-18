# Product Requirements Document: KAAVIS
### Voice-First Personal AI Agent
**Status:** Draft v0.1
**Owner:** You
**Last updated:** July 18, 2026
---
## 1. Vision
KAAVIS is a voice-first personal AI agent that acts as a unified layer over your personal digital life — email, documents, notes, and browsing — controllable primarily by voice, with a companion UI that lives on your local machine. Instead of switching between Mail, Word, Notes, Obsidian, Safari, and Google Docs, you talk to KAAVIS and it reads, writes, edits, organizes, and reasons across all of them on your behalf.
The core bet: most personal computing friction is *context-switching and manual file wrangling*. KAAVIS removes that by becoming the single interface — you speak intent, it executes across systems, and it shows its work in a local UI you can inspect, undo, and correct.
## 2. Goals
- **G1 — Voice as primary input.** Every core action (read, draft, edit, search, organize, summarize) is achievable by voice, with a visible/editable UI as the confirmation and audit layer, not the primary input method.
- **G2 — Unified personal context.** KAAVIS understands the relationships between your email, docs, notes, and vault (e.g., "the proposal I emailed Raj about" resolves to the right Word file).
- **G3 — Local-first control surface.** The UI and file operations run on your local machine, giving you a persistent, inspectable record of every action KAAVIS takes — not a black-box cloud agent.
- **G4 — Safe-by-default file and account access.** Read access is broad and low-friction; write/delete/send actions are scoped, logged, reversible where possible, and confirmed for anything irreversible or externally visible.
- **G5 — Cross-device continuity.** Actions initiated by voice on iPad (Safari, Google Docs, Notes) and actions on the local machine (Obsidian vault, Word files, directories) feel like one system, not two.
## 3. Non-Goals (v1)
- Not building a general-purpose OS-level automation framework — scope is the specific integrations listed below.
- Not attempting real-time collaborative multi-user editing.
- Not replacing Siri/Shortcuts at the OS level — KAAVIS orchestrates *through* available APIs and automation surfaces, not by replacing iOS/macOS system agents.
- Not building its own email/document storage — KAAVIS always operates on your existing accounts and files, never forks a parallel data store.
- No autonomous background operation in v1 (e.g., auto-replying to emails unsupervised) — all writes are agent-proposed, user-confirmed, until trust is established.
## 4. Personas & Context
Single-user, personal system. You are simultaneously the product owner, the primary user, and the admin. This simplifies auth (no multi-tenant concerns) but raises the bar on personal data safety, since a single compromised credential or bad agent action touches your entire digital life.
## 5. System Architecture (High Level)
```
┌─────────────────────────────┐        ┌──────────────────────────┐
│   iPad (voice + touch)      │        │  Local Machine (Windows 10)│
│  - Voice capture            │        │  - KAAVIS UI (control     │
│  - Safari automation        │        │    surface, file/diff     │
│  - Notes / Google Docs      │        │    viewer, action log)    │
│    surface                  │        │  - Local file/dir agent   │
│  - Relay point for Apple    │        │  - Obsidian vault access  │
│    Notes content (see §6)   │        │  - Word file access       │
└──────────────┬───────────────┘        └──────────────┬────────────┘
     ┌───────────────────────────────────────────────────┐
     │              KAAVIS Orchestrator (core)             │
     │  - Voice intent parsing                             │
     │  - Task planning / tool routing                     │
     │  - Cross-source context resolution                  │
     │  - Permission & confirmation gate                   │
     │  - Action log / undo ledger                         │
     └───────────────────────────────────────────────────┘
               │              │              │
       ┌───────▼────┐  ┌──────▼──────┐  ┌────▼────────┐
       │ Email       │  │ Google      │  │ Apple ID /   │
       │ (Gmail/     │  │ (Drive,     │  │ Notes/       │
       │ IMAP)       │  │ Docs API)   │  │ iCloud       │
       └─────────────┘  └─────────────┘  └──────────────┘
```
**Key architectural decision:** the *orchestrator and permission gate* run once, centrally (local machine is the natural home given it's always-on and has the most privileged file access). The iPad is a **voice + rendering client**, not a second brain — it sends voice/intent to the orchestrator and displays results, rather than running independent logic. This avoids state-sync bugs between two "smart" endpoints.
## 6. Integration Surfaces — Detail, Feasibility, and Constraints
Each integration has a real-world access method and real constraints. These are noted explicitly because they materially affect scope and sequencing.
| Surface | Access method | Constraints to design around |
|---|---|---|
| **Email** | Gmail API (OAuth) or IMAP/SMTP | Full read/send scopes are sensitive — request narrowest scope needed per feature (read-only first, send later). Rate limits, thread-reconstruction complexity. |
| **Word files** | Local filesystem + Office Open XML (python-docx or similar); optionally COM automation (pywin32) against a running Word install for live edits, or Microsoft Graph API if using OneDrive | Local files: direct read/write, no auth needed. COM automation gives "edit the doc that's actually open" behavior but ties KAAVIS to Word being installed/running. Cloud-synced (OneDrive/SharePoint) files need Graph API + separate auth. Track which files are local-only vs. cloud-synced. |
| **Google account (Drive/Docs)** | Google Drive API + Docs API (OAuth) | Google Docs' native format isn't a flat file — edits must go through the Docs API, not treated as a generic file. Separate scope from Gmail/Drive. Platform-agnostic — works the same from Windows as from iPad. |
| **Apple ID / iCloud** | **No general public API, and no AppleScript on Windows.** Real options: (a) **iCloud for Windows** app — syncs iCloud Drive, Photos, Bookmarks, Mail/Contacts/Calendar via Outlook, but **does not sync the Notes app**; (b) iPad-side **Shortcuts automation** as the only way to actually touch Notes content, relaying results to Windows via a file dropped in iCloud Drive (which *is* synced) or emailed to yourself | This is the biggest constraint in the whole system, and Windows makes it tighter than macOS would. Treat "Apple ID access" as: **iCloud Drive files (native sync to Windows)** + **Notes content (iPad-only, via Shortcuts, relayed to Windows through iCloud Drive as an intermediate file)** — there is no direct Windows-to-Notes path. |
| **Obsidian vault** | Direct filesystem access (vault = plain folder of Markdown files) | Easiest integration — no API needed, just file I/O. Must respect Obsidian's own file-watching (avoid write conflicts while Obsidian has the file open). |
| **Safari (iPad)** | iOS Shortcuts app (Safari actions), or Accessibility-based automation | No official Safari scripting API on iOS. Realistically limited to: opening URLs, reading page content via Shortcuts' "Get Contents of Web Page," triggering Reader mode. Full "browse and click for me" agentic control is **not feasible on iPad in v1** without jailbreak-adjacent tooling, which is out of scope. |
| **Google Docs (iPad)** | Google Docs API (same as above), or the Docs app via Shortcuts if API insufficient | Prefer API — works identically whether triggered from iPad or local machine. |
| **Local machine files/directories** | Direct filesystem access via Windows file I/O, sandboxed to explicitly granted directories (app-level path allow-list, since Windows has no macOS-style per-folder permission prompts) | Full-disk access is a large ask and Windows won't gate it for you the way macOS does — KAAVIS itself must enforce an explicit **allow-list of directories** (e.g., `Documents\`, `Obsidian Vault\`) rather than relying on the OS to prompt. |
**Important scoping call-out:** "access to my Apple ID" as stated in the request isn't a single grantable permission — Apple doesn't expose one, and on Windows there's no AppleScript fallback either. The PRD should tell the build team explicitly: *Apple ID access = iCloud Drive files (native sync via iCloud for Windows) + Notes content (only reachable through iPad-side Shortcuts, relayed back to Windows via a synced file), not a general account API and not something the Windows machine can query directly.* This should be validated with the user before implementation starts — if Notes access turns out to be a must-have on day one rather than a nice-to-have, it may be worth reconsidering whether Notes or Obsidian is the "true" notes source of record for KAAVIS.
## 7. Core User Stories (v1 scope)
1. *"KAAVIS, summarize my unread emails from this morning and flag anything urgent."*
2. *"Find the Word doc I was working on about the Q3 roadmap and read me the open questions section."*
3. *"Add a note to my 'Groceries' note in Apple Notes: oat milk, batteries."*
4. *"Create a new note in my Obsidian vault under Projects/KAAVIS summarizing this conversation."*
5. *"Open the client proposal in Google Docs and add a paragraph about the revised timeline."*
6. *"What did I write in my journal about the Denver trip?"* (cross-source: Obsidian + Notes)
7. *"Draft a reply to Priya's email but don't send it — show me first."*
8. *"Rename all the screenshots in my Downloads folder by the date they were taken."*
## 8. UX Principles
- **Voice initiates, UI confirms.** Every write/send/delete action surfaces a diff or preview in the local UI before committing, unless the user has explicitly pre-approved that action class ("always save notes without asking").
- **Always show the plan before irreversible actions.** For multi-step tasks ("clean up my Downloads folder"), KAAVIS states its plan aloud/on-screen and waits for a go-ahead before executing more than one destructive step.
- **One source of truth for state.** The local UI's action log is the canonical record of everything KAAVIS has done, across both devices.
- **Graceful voice fallback.** If a voice command is ambiguous, KAAVIS asks one clarifying question rather than guessing on anything write-related; for read-only actions it can proceed on best interpretation and say what it assumed.
## 9. Edge Cases & Failure Modes
### Voice & intent
- Ambiguous referents ("delete that file" — which file, from which of 5 recently mentioned?) → require disambiguation before any destructive action.
- Similar-sounding names/commands (e.g., "Notes" vs "Note" vs a file literally named "notes.docx") → confirm target explicitly when confidence is low.
- Background noise / partial capture leading to a mis-transcribed command that changes meaning (e.g., "send" vs "send it") → never execute a send/delete/overwrite on a single-pass transcription without a confirmation step.
- Multi-turn context drift over a long session (KAAVIS "forgetting" what "it" refers to) → explicit session context window with visible current-context indicator in the UI.
### Cross-account / cross-source conflicts
- Same file open and being edited in both Obsidian and by KAAVIS simultaneously → file-lock/watch detection; refuse to write if externally modified since last read, surface a merge/conflict view instead of silently overwriting.
- Google Doc edited via API while user is actively editing it in the browser → Docs API's own revision handling should be used; warn user if a live-edit collision is detected.
- Duplicate/near-duplicate content across Notes, Obsidian, and Word (e.g., three "meeting notes" for the same meeting) → KAAVIS should present candidates and ask rather than auto-picking one.
### Permissions & security
- Partial auth failure mid-task (e.g., Gmail token expires mid-summarization across 20 emails) → task should fail gracefully with a clear "processed 12 of 20, re-auth needed" state, not a silent partial result presented as complete.
- Scope creep — a feature requesting broader access than it needs (e.g., email *drafting* only needing compose scope, not full mailbox access) → each integration requests the narrowest OAuth scope for its current feature set; broader scopes are opt-in per feature, not granted up front.
- Local file access outside allow-listed directories (accidental or requested) → hard block, not just a warning; require explicit re-scoping through the UI to add a new directory.
- Sensitive content handling (financial info in emails, personal info in notes) → these should never be sent to any third-party model/service beyond what's necessary for the requested action, and never logged in plaintext in the action log beyond what's needed for undo.
### Reversibility
- Every write action logged with enough information to undo it (previous file version, previous note content, "unsent" draft state) where technically possible.
- Email send is **not reversible** once sent — this is the one action class that should always require explicit confirmation, no "always allow" override.
- File deletion should move to a KAAVIS-managed trash/archive folder rather than permanent deletion, for at least 30 days.
### Platform/device edge cases
- iPad Safari automation silently failing because a page's structure doesn't match what Shortcuts expects → detect failure explicitly (don't assume success), report back rather than hanging.
- Local machine asleep/offline when a voice command comes from iPad → queue the request, tell the user it's queued, execute when the machine reconnects, rather than failing silently.
- iCloud sync lag causing KAAVIS to act on a stale version of a Note or file → where possible, force a sync check before reading/writing iCloud-backed content.
### Trust & correction
- Wrong action taken → one-command undo ("KAAVIS, undo that") should work for the majority of write actions via the action ledger.
- Repeated misunderstanding of the same type of command → KAAVIS should track and surface patterns ("I've misread this command 3 times — want to rephrase how you usually ask for this?") rather than the user silently working around it forever.
## 10. Permission Model (Summary)
- **Default posture:** read access broad within granted scopes; write access explicit and confirmed; send/delete access always confirmed regardless of prior settings.
- **Per-integration scoping**, not one global "connect everything" toggle — user can revoke Gmail without affecting Obsidian access, etc.
- **Directory allow-list** for local filesystem access, editable in the UI at any time.
- **Action log** is visible, searchable, and exportable — this is the accountability layer that makes broad access tolerable.
## 11. Success Metrics (v1)
- % of voice commands resolved without a clarifying question needed.
- % of write actions that required a correction/undo (target: trending down over time as trust calibration improves).
- Time-to-complete for the 8 core user stories in Section 7, voice vs. manual baseline.
- Zero unintended sends/deletes across a 30-day dogfooding period.
## 12. Suggested Build Sequence
1. **Local file + Obsidian vault integration** — no external API, fastest to build, proves the local UI + action log architecture.
2. **Word file read/write** (local files first, OneDrive/Graph API later).
3. **Google account (Drive + Docs API)** — read first, then write.
4. **Gmail (read-only)** — summarization/search before any send capability.
5. **Gmail send/draft** — gated behind the confirmation UI proven in step 1–4.
6. **Apple Notes via iPad Shortcuts + iCloud Drive relay** — scoped explicitly as its own integration, not bundled into "Apple ID." Build the iCloud Drive relay file format first (simple: a synced Markdown/text file Shortcuts writes to and Windows watches), since there's no direct API path.
7. **iPad Safari via Shortcuts** — read/open-URL actions only in v1; revisit deeper automation later if a viable API/tooling path emerges.
## 13. Open Questions for You
- For email: Gmail specifically, or also other providers (Outlook/iCloud Mail)?
- Given Notes is only reachable via the iPad relay described in Section 6, is Apple Notes still worth building for v1, or should Obsidian serve as the single notes source of truth and Notes get deprioritized?
- Should KAAVIS ever act without a per-action confirmation (e.g., a fully trusted "auto-file my notes" mode), or is confirm-always the permanent default?
- Voice capture: always-listening wake word, or push-to-talk only? This has real privacy implications given the breadth of data access.
- Where should the action log/undo ledger persist — local only, or also backed up (and if so, encrypted how)?
---
*This is a v0.1 draft meant to be argued with — Section 6 in particular (Apple ID/Safari constraints) should be validated against current Apple/Google API terms before any build work starts, since platform policies shift.*
