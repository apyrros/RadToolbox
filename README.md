# RadToolbox.ahk
AutoHotkey v1 script that adds AI-assisted tools to Nuance PowerScribe 360 and Epic Hyperspace style text boxes using only keystrokes and clipboard operations (no vendor SDK).

## Features
- Generate Impression (few-shot aware, optional numbering) and color the AI output in the report.
- Check Report for Errors with a clickable error terminal.
- Restore Dictation (local undo history buffer).
- Prompt Manager with editable prompts, custom prompt hotkeys, and in-place text transforms colored as AI output.
- Pull Indication from Epic / clipboard and clean it.
- AI Notes helper for Epic with copy/insert options.
- Preferences dialog: provider/endpoint/model, API key, hotkeys, AI text color presets, Epic Notes toggle.
- Startup EULA, portable config under `%AppData%\AI_Tools_Demo\settings.ini`.

## Requirements
- AutoHotkey v1 (32-bit recommended for the embedded JSON library).
- Access to PowerScribe 360 (window title should match `PowerScribe 360 | Reporting`) or Epic Hyperspace.
- API key for OpenAI or Azure OpenAI; Ollama supported without a key when running locally.

## Setup
1) Place `RadToolbox.ahk` on the workstation with AutoHotkey v1 installed.  
2) Run the script; accept the startup disclaimer.  
3) Open Preferences to set Provider, Endpoint, Model, and API Key (if needed).  
4) Verify the PowerScribe window title matches the default; adjust `PowerScribeWindowTitle` in the script if your site differs.

## Usage
- Launch `RadToolbox.ahk`; the main GUI stays on top for quick access.
- Use toolbar buttons or menus for actions; hotkeys are enabled by default (see below).
- For custom prompts: highlight text in the report, choose a prompt from the Custom Prompts menu or press its hotkey; the transformed text is inserted in place and colored with the AI text color.
- For impression generation: click inside the report, run Generate Impression; the script inserts the IMPRESSION section and applies AI coloring.

### Default hotkeys
- `Ctrl+Alt+I` Generate Impression
- `Ctrl+Alt+E` Check Report for Errors
- `Ctrl+Alt+R` Restore Dictation
- `Ctrl+Alt+P` Custom Prompt selector
- `Ctrl+Alt+N` AI Notes (Epic)
Custom prompt-specific hotkeys can be assigned per prompt in the Prompt Manager (must include a modifier or be an F-key).

### Prompt Manager
- Add, edit, or delete prompts; each prompt can have an optional hotkey.
- Prompts appear under the Custom Prompts menu; selecting one applies it to the current selection or caret position.
- In-place replacements inside PowerScribe are colored with the configured AI text color (`Preferences -> AI text color`).

### Few-shot archive
- When you add reports via the archive helper, they are stored as JSONL under `%AppData%\AI_Tools_Demo\report_archive`.
- Impression generation reuses the most recent examples per exam type (configurable limit).

## Configuration storage
- Settings: `%AppData%\AI_Tools_Demo\settings.ini`
- Report archive: `%AppData%\AI_Tools_Demo\report_archive\*.jsonl`
- Single-file design; no external includes.

## Troubleshooting
- If the script cannot find the PowerScribe text box, ensure the active window title contains `PowerScribe 360 | Reporting` and the box is editable.
- If API calls fail, verify endpoint, model, and API key in Preferences; Ollama must be reachable at `http://localhost:11434` by default.
- If custom prompts fail to color text, ensure the selection is inside the PowerScribe report box; coloring is best-effort and non-fatal.
- If JSON parse errors appear when loading archives, they are skipped automatically; re-add reports using the current script if needed.

## Privacy, PHI, and HIPAA
- This script sends text to external AI endpoints (OpenAI, Azure OpenAI, or Ollama). **Do not process real PHI unless you are using a HIPAA-eligible, properly configured, and BAA-covered endpoint approved by your organization.**
- Ensure outbound network policies, logging, and audit requirements are met before enabling API calls in a clinical environment.
- Clipboard contents are temporarily used for operations; the script attempts to restore the clipboard after each action, but users should avoid copying PHI unless policies permit.
- Users are responsible for verifying all AI-generated content before it enters the medical record; AI output may be incomplete or inaccurate.

## Disclaimer
- Provided as-is with no warranty; not cleared for clinical use by any regulatory authority.
- The user is solely responsible for compliance with institutional policies, HIPAA, and applicable laws.
- Always review and edit AI output; do not rely on it for clinical decision-making without appropriate validation.
