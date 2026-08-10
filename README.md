# Mute Browse Mode

An NVDA add-on for NVDA 2026.1 that stops NVDA announcing a browse mode document
every time one loads or is entered, and replaces that announcement in a web browser
with a short summary of the page.

NVDA normally speaks the document name, the word "document" and the first line of
the buffer whenever a web page finishes loading in Chrome, whenever you switch to a
browser window, and whenever you open a message in any version of Microsoft Outlook.
This add-on lets you silence that.

## Download

Grab the latest `.nvda-addon` file from the
[Releases](../../releases/latest) page. Open it to install, then restart NVDA.

## The settings

Two combo boxes are added to NVDA's Browse Mode settings
(NVDA menu → Preferences → Settings → Browse Mode).

### Mute browse mode

| Choice | Behaviour |
| --- | --- |
| Silence all browsing | NVDA stays silent while the document is loading and when it has finished loading. "Loading document..." is not spoken either. |
| Play tones | Speech is silenced the same way, and three quick ascending tones (A4, D5, G5) play when the document is ready. |
| Normal | NVDA behaves exactly as if the add-on were not installed. This is the default. |

### Announce a page summary when a page has loaded

What happens in a web browser once a page has finished loading, in place of the
announcement the combo box above silenced.

| Choice | Behaviour |
| --- | --- |
| Speak the summary | NVDA says "Loading complete", then "Page has 8 regions, 57 headings and 196 links". This is the default. |
| Play tones | The three ascending tones instead of the summary. If the combo box above is also set to tones, you still only hear them once. |
| Normal | No summary. |

### Say "loading complete" before the page summary

A check box under the two combo boxes, ticked by default. Untick it and the summary is
just "Page has 8 regions, 57 headings and 196 links".

Both combo boxes also have a cycle command in the Input Gestures dialog under
"Mute Browse Mode", with no gesture assigned by default.

## Switching windows

Nothing is silenced while NVDA is announcing a window you have just switched to. The
window title and the control the focus landed on are read right to the end, and only
then does the silencing resume. The browse mode document itself stays silent
throughout, which is the whole point of the add-on.

"Read right to the end" is the synthesiser's word for it rather than a guess at how
long a title takes: a `speech.commands.CallbackCommand` is appended to each utterance
of the announcement, and NVDA runs it when speech actually reaches that point. Pressing
a key ends the wait early, and a ceiling of eight seconds ends it whatever happens, so
a cancelled announcement can never leave the add-on switched off.

Each announcement is numbered and every callback carries its number. It has to: NVDA
hands those callbacks back through the event queue — `_onSynthIndexReached` queues
`_handleIndex`, which runs the callback — so one belonging to the window you have just
left routinely arrives *after* the switch to the next window has begun. Counting it
against the new announcement drove the pending count negative, ended the announcement
on the spot, and let the deadline gate eat the rest of the new window's title. That is
why switching out of Outlook, the one place the add-on speaks an utterance of its own,
was where the half-read title showed up. An announcement also stays open for 250 ms
after its last utterance, so a gap between two parts of one cannot close it either.

## Switching to Microsoft Outlook

Arriving in Outlook from another program answers with a brief description of where you
have landed — the field, the focus, what is in it, and where it sits:

> Inbox, list, Josh Kennedy build 1.4, list item, 3 of 57

Where the focus is a control holding a choice, such as a combo box, the selected item
is included. This **replaces** NVDA's own report rather than adding to it, so nothing
is said twice: NVDA's report still runs, silently, so braille and NVDA's property
caches are untouched. It only fires when Outlook was not already the foreground
program, and only when the mute browse mode combo box is not on Normal.

## The Outlook message body

Tabbing through a new message goes To, Cc, Subject, body. NVDA names the first three,
but the body has no name to read, so nothing tells you that you have reached the part
you type into. When the focus lands there NVDA now says:

> You are now in the message body, type a message.

Said after NVDA's own announcement of the field, not before — a message body that is a
browse mode document calls `speech.cancelSpeech()` on its way into focus mode, so
anything said first would be cut off by it.

Only for a body you can type into. The address and subject fields are editable text
too, so the test is not "is this editable": the body is the one that takes more than
one line, or is a whole document rather than a field, or lives in a window Outlook only
ever puts a message body in (`_WwG`, `Internet Explorer_Server`, `RichEdit20W`). A
read-only body, such as the reading pane, is left alone, because the announcement
invites you to type. Tabbing away and back says it again; the same body raising a
second focus event does not.

## Where the summary happens

Only when **Microsoft Outlook is not the program you are in**. That is confirmed
silently on every single page load, from the foreground window and from the focus;
nothing about the check is ever spoken. Outlook is the one application the add-on
exists to keep quiet, so a message opening there stays as quiet as it was.

It is announced in Chromium browsers (Chrome, Edge, Brave, Opera, Vivaldi and forks),
in Gecko browsers (Firefox and forks) and in Electron applications. The new Outlook for
Windows is a WebView2 application, so it looks exactly like Chromium from the outside;
it is matched as Outlook first and stays silent.

Regions are NVDA's landmarks, and all three counts come from the same
`_iterNodesByType` search the Elements List uses. A page with a huge number of any one
element is not counted all the way to the end — the count stops at 1500 or after 1.5
seconds, whichever comes first, and the summary says "over 1500 links" — so that
counting can never hold NVDA up.

The summary is also skipped for a page that loaded in a background tab, for a page you
started using before it finished loading (any keypress cancels it), and while say all
is running.

## Control+F in Outlook

Control+F is handed straight to whichever program you are in and is never claimed by
NVDA. In Outlook that forwards the message you are reading, which is what control+F
means in Outlook. In a browser it opens the browser's find bar.

NVDA's find stays on **NVDA+control+F**, which is where NVDA itself binds it, with
NVDA+F3 and NVDA+shift+F3 for the next and previous match.

The add-on binds control+F rather than leaving it unbound because a global plugin
script is the first thing NVDA looks for, ahead of the browse mode document and
anything else that might claim the key. The script does nothing but hand the keystroke
on, and NVDA ignores the keys it injects itself, so it cannot come back round.

## In Outlook and Chromium browsers

Neither of the two silencing choices stops at the document announcement in Microsoft
Outlook or in Chromium based browsers. There, NVDA also stops speaking:

- **window titles**, so switching to Outlook no longer reads the title of the message
  list window and opening a message no longer reads the title of the message window;
- **dialogs**, so opening an Outlook message no longer says the word "dialog". A
  message opens inside a dialog, and NVDA announces it while walking down the focus
  ancestors, which happens before the message document exists;
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
NVDA+t for the title, NVDA+tab for the focus and NVDA+b for a whole dialog all use
`OutputReason.QUERY`, which is never suppressed.

## Building

```bash
python build.py
```

That writes `muteBrowseMode-1.5.nvda-addon` next to `build.py`. Open it to install,
or drag it onto NVDA.

## How it works

`speech.speech.speak` is wrapped with a gate, and the gate is held shut across the
four separate points in NVDA where a document gets announced.

There are two of them. While a hooked call is on the stack a **depth counter** silences
everything, and that is the one that actually swallows the document announcement —
NVDA speaks all of it inline, so being inside the call is enough. Once the call
returns only the **deadline gate** is left, and that one stands down while a new window
is being announced. `eventHandler.doPreGainFocus` runs `foreground`, then
`focusEntered` for the window title, then `event_treeInterceptor_gainFocus` for the
document, then `gainFocus` for the control; splitting the gate this way is what lets
the first and last through while still silencing the middle one. They are genuinely
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
  "Refreshed" on a reload, and the moment the chime and the page summary belong to.

Gating `speak` rather than each individual announcement means every route into the
synthesiser is covered: `speakTextInfo`, `speakObject`, `speakMessage` and
`ui.message` all funnel through it.

The page summary rides on the same two load hooks, but outside the mute mode check, so
that it has its own setting rather than being switched off by the other one.
`loadBuffer` marks the buffer as owed a summary, `_loadBufferDone` schedules it 400 ms
later — long enough for the tail of NVDA's own announcement, which is queued onto the
main queue, to have come and gone — and it is spoken inside a bypass that lets the
add-on's own speech past its own gate. A buffer that finishes loading empty is one NVDA
is still waiting on, and it stays armed for `event_documentLoadComplete`, which is
where NVDA reports such a document instead.

On top of the gate, `speech.speakObject` drops window, pane, frame, dialog, document,
application and alert roles in Outlook and Chromium windows, and
`NVDAObject.event_liveRegionChange`, `behaviors.Notification.event_alert` /
`event_show` and `IAccessible.event_alert` are silenced there. Those calls still run,
so braille and NVDA's caches are unaffected — only the speech is dropped.

The dialog is dropped by role rather than gated because there is nothing to gate it
with yet. `eventHandler.doPreGainFocus` fires `focusEntered` for every new focus
ancestor — which is where `NVDAObject.event_focusEntered` calls
`speakObject(reason=FOCUSENTERED)` on the dialog an Outlook message opens in — and only
then reaches `event_treeInterceptor_gainFocus`. The word "dialog" is therefore spoken
before the browse mode document exists at all.
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
  still read the page after it loads, and the page summary is skipped so it does not
  talk over it. Turn it off if you want the silence.

## License

GNU General Public License, version 2.
