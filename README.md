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

## In Outlook and Chromium browsers

Neither of the two silencing choices stops at the document announcement in Microsoft
Outlook or in Chromium based browsers. There, NVDA also stops speaking:

- **window titles**, so switching to Outlook no longer reads the title of the message
  list window and opening a message no longer reads the title of the message window;
- **document titles**, so a browser tab no longer announces the page title;
- **toasts and notification balloons**, such as a browser's download and pop-up
  messages;
- **live region "flash" messages**, which is how web pages announce things like
  "Message sent".

Chromium browsers are recognised by the window classes Chromium creates
(`Chrome_WidgetWin_0`, `Chrome_WidgetWin_1`, `Chrome_RenderWidgetHostHWND`) as well as
by executable name, so Chrome, Edge, Brave, Opera, Vivaldi and forks of them are all
covered without the add-on needing to know about them, as is the new Outlook for
Windows, which is a WebView2 application.

Only announcements NVDA volunteers are dropped. Anything you ask for still answers:
NVDA+t for the title and NVDA+tab for the focus both use `OutputReason.QUERY`, which
is never suppressed.

## Building

```bash
python build.py
```

That writes `muteBrowseMode-1.1.nvda-addon` next to `build.py`. Open it to install,
or drag it onto NVDA.

## How it works

`speech.speech.speak` is wrapped with a gate, and the gate is held open across the
four separate points in NVDA where a document gets announced. They are genuinely
separate: silencing the document load does not silence the Outlook message, because
NVDA does not read an Outlook message as part of a load.

- `browseMode.BrowseModeDocumentTreeInterceptor.event_gainFocus` — where NVDA calls
  `speakTextInfo(reason=FOCUS)` on the line the focus landed on. **This is what reads
  the first line of an Outlook message.** It is gated only when
  `_enteringFromOutside` is set, which is NVDA's own flag for "focus arrived from
  outside this document", so tabbing and arrowing around inside a document keep
  speaking normally.
- `browseMode.BrowseModeDocumentTreeInterceptor.event_treeInterceptor_gainFocus` —
  where NVDA calls `speakObject` on the document root (the name plus "document") and
  `speakTextInfo` on the current line. This fires the first time a document is
  entered.
- `virtualBuffers.VirtualBuffer.loadBuffer` — the start of a load, which is what
  silences "Loading document...".
- `virtualBuffers.VirtualBuffer._loadBufferDone` — the end of a load, which speaks
  "Refreshed" on a reload, and the moment the chime belongs to.

Gating `speak` rather than each individual announcement means every route into the
synthesiser is covered: `speakTextInfo`, `speakObject`, `speakMessage` and
`ui.message` all funnel through it.

On top of the gate, `speech.speakObject` drops window, pane, frame, document,
application and alert roles in Outlook and Chromium windows, and
`NVDAObject.event_liveRegionChange`, `behaviors.Notification.event_alert` /
`event_show` and `IAccessible.event_alert` are silenced there. Those calls still run,
so braille and NVDA's caches are unaffected — only the speech is dropped.
`OutputReason.ONLYCACHE` is never suppressed, because browse mode relies on it to keep
its property cache honest.

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
