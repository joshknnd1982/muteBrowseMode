# Mute Browse Mode

An NVDA add-on for NVDA 2026.1 that stops NVDA announcing a browse mode document
every time one loads or is entered.

NVDA normally speaks the document name, the word "document" and the first line of
the buffer whenever a web page finishes loading in Chrome, whenever you switch to a
browser window, and whenever you open a message in any version of Microsoft Outlook.
This add-on lets you silence that.

## Download

Grab the latest `.nvda-addon` file from the
[Releases](../../releases/latest) page. Open it to install, then restart NVDA.

## The setting

A **Mute browse mode** combo box is added to NVDA's Speech settings
(NVDA menu → Preferences → Settings → Speech), with three choices:

| Choice | Behaviour |
| --- | --- |
| Silence all browsing | NVDA stays silent while the document is loading and when it has finished loading. "Loading document..." is not spoken either. |
| Play tones | Speech is silenced the same way, and three quick ascending tones (A4, D5, G5) play when the document is ready. |
| Normal | NVDA behaves exactly as if the add-on were not installed. This is the default. |

There is also a "cycle mute browse mode" command in the Input Gestures dialog under
"Mute Browse Mode", with no gesture assigned by default.

## Building

```bash
python build.py
```

That writes `muteBrowseMode-1.0.nvda-addon` next to `build.py`. Open it to install,
or drag it onto NVDA.

## How it works

`speech.speech.speak` is wrapped with a gate, and the gate is held open across three
points in NVDA where the document announcement happens:

- `browseMode.BrowseModeDocumentTreeInterceptor.event_treeInterceptor_gainFocus` —
  where NVDA calls `speakObject` on the document root (the name plus "document") and
  `speakTextInfo` on the current line (the first line). This covers web pages, and
  also Outlook and Word message bodies, which are browse mode documents but not
  virtual buffers.
- `virtualBuffers.VirtualBuffer.loadBuffer` — the start of a load, which is what
  silences "Loading document...".
- `virtualBuffers.VirtualBuffer._loadBufferDone` — the end of a load, and the moment
  the chime belongs to.

Gating `speak` rather than each individual announcement means every route into the
synthesiser is covered: `speakTextInfo`, `speakObject`, `speakMessage` and
`ui.message` all funnel through it.

The gate is a deadline, not a counter, so an exception can never leave NVDA
permanently mute — the worst case is a few seconds of silence that expires on its
own. Any input gesture closes it immediately via the `decide_executeGesture`
extension point, so pressing a key always restores normal speech at once. Say all is
explicitly let through, so "automatically say all on page load" keeps working.

Braille is deliberately untouched.

## Notes

- If Browse Mode settings → "Audio indication of focus and browse modes" is on, you
  will still hear that short sound when entering a document, because it is a sound
  rather than speech. Turn it off and NVDA speaks "browse mode" instead, which this
  add-on does silence.
- If Browse Mode settings → "Automatically say all on page load" is on, NVDA will
  still read the page after it loads. Turn it off if you want the silence.

## License

GNU General Public License, version 2.
