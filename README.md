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

### Links are on their own line

A check box, unticked by default. When ticked, an Outlook message puts each link,
button and other control on a line of its own, so down arrow reaches them one at a
time.

Outlook renders messages two completely different ways, and each needs its own answer.

**A message rendered as a web document** — Outlook on the web, the new Outlook for
Windows, and messages classic Outlook shows through `Internet Explorer_Server` — is a
virtual buffer, so NVDA's own **Use screen layout (when supported)** already covers it.
That setting is read in exactly one place, `VirtualBufferTextInfo._getLineOffsets`, and
passed straight to `VBuf_getLineOffsets`, which breaks the line at any node with
children when it is off. The add-on wraps that one method and asks for the line as if
screen layout were off, but only for an Outlook message, so the setting your browsers
use is never touched. If NVDA's screen layout is already off, this has nothing to add
and stays out of the way.

**A message rendered by Word** — which is what classic Outlook does by default — is not
a virtual buffer at all, and NVDA has nothing equivalent: the base implementation of
`script_toggleScreenLayout` answers "Not supported in this document." Its lines are
Word's own, so a link in the middle of a sentence is read as part of that sentence and
down arrow steps straight over it.

Redefining the line unit would reach far too much — NVDA uses `UNIT_LINE` for braille,
for reporting the line the focus lands on, and elsewhere. So only
`CursorManager.script_moveByLine_forward` and `script_moveByLine_back` are wrapped, and
only for this one kind of document. They walk the line in segments split where a control
starts and ends, so the link is on its own and the words either side of it are too:

> "comment this is a link" → down arrow → "comment " → down arrow → "this is a link"

Everything else about the document is left exactly as NVDA has it, and any difficulty at
all falls straight back to NVDA's own line movement. `resumeSayAllMode` is carried onto
the wrappers, so say all still resumes from an arrow key.

### Control+F opens NVDA's find in a web browser

A check box, ticked by default. Control+F in a web browser opens NVDA's find — in a
browse mode buffer or in an edit box — instead of the browser's own find bar.

### Bring up NVDA screen reader find when not in Outlook

A check box, unticked by default. Widens the one above from web browsers to everywhere
except Microsoft Outlook, where control+F stays Forward. Where both are ticked this one
wins. See [Control+F opens NVDA's find](#controlf-opens-nvdas-find-in-a-browser) for
what counts as somewhere NVDA can search.

Both combo boxes also have a cycle command in the Input Gestures dialog under
"Mute Browse Mode", with no gesture assigned by default.

## The Outlook spelling checker

When the F7 spelling window takes the focus, NVDA says the word the checker is asking
about and then spells it out, a fifth more slowly than your usual rate, with the rate
back to normal straight afterwards.

The slowdown is a `speech.commands.RateCommand(multiplier=0.8)` inside that one
announcement, followed by a plain `RateCommand()` to return to the configured value. It
is a synth parameter carried in the speech sequence, not a change to the synthesiser's
settings, so it lasts exactly as long as the spelling does and cannot leak anywhere
else — not even if the announcement is interrupted half way through.

Finding the word:

- The dialog's box is a Word editing surface with window class `_WwN`, which NVDA has a
  class of its own for (`WordDocument_WwN`). The focus often lands on the suggestions
  list rather than that box, so it is looked up by window class in the same thread with
  `NVDAHelper.localLib.findWindowWithClassInThread`, the same call NVDA uses to find the
  document behind such a box. It only counts when `winUser.getAncestor(..., GA_ROOT)`
  says it belongs to the same top-level window as the focus, so a Word surface in some
  other Outlook dialog is never mistaken for it.
- `WordDocument_WwN` points `WinwordSelectionObject` at the application's active pane,
  and Word selects the error in the message itself as it steps through — so
  `makeTextInfo(POSITION_SELECTION)` on that object *is* the misspelled word. The word
  under the cursor is the fallback.
- Whatever comes back must look like a single word (non-empty, no whitespace, at most
  60 characters) or it is discarded and nothing is said. The box holds the whole
  sentence the error is in, and reading that out would be worse than silence.

The word is remembered with its window, so the dialog moving to the next error announces
that one while a repeated focus event on the same one does not. Leaving the dialog
clears it, so coming back announces again.

This is not tied to the mute browse mode combo box: it adds something rather than
silencing something, and works whatever that setting is on.

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

### The window title itself

A browser window's title is read in full on a switch, even though browser window titles
are dropped by role the rest of the time. Switching between two browser windows is what
needs it: the alt+tab switcher names each window as you cycle, `event_foreground` calls
`speech.cancelSpeech()` the moment one activates, and with the post-switch title also
suppressed the user was left with only the fragment the switcher got through. So
`_shouldDropObjectSpeech` releases the window-title roles (`WINDOW`, `PANE`, `FRAME`,
`INTERNALFRAME`, `APPLICATION`) while a foreground announcement is running, outside
Outlook. `DIALOG` and `DOCUMENT` stay dropped — the dialog is the word an opening
Outlook message says, and the document is the page title the add-on exists to silence.

`speech.cancelSpeech` is also wrapped, and does nothing while a hooked document call is
on the stack during a foreground announcement.
`BrowseModeDocumentTreeInterceptor.event_gainFocus` cancels speech when the focus lands
somewhere that forces focus mode, on the reasoning that a focus change should stop the
page being read aloud; arriving from another window there is no page being read aloud
yet, only the title, and cancelling that is what cut it off. Every other cancel in NVDA,
including the one the foreground change itself makes, is untouched.

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

Only for the message body itself. Nearly everything on a message form is editable, so
"is this editable" is not the test, and neither is "does it take more than one line" —
Outlook's recipient fields wrap, so they are multiline rich edit controls too. The body
is identified positively instead, by any of:

- **`obj.isReadonlyViewer` on a `_WwG` window** — NVDA's `appModules/outlook.py` puts
  `isReadonlyViewer` on the message body object (`BaseOutlookWordDocument`), and its
  value says whether the message is being written or read. It is not enough on its own:
  Word hands the same kind of object to its dialogs, and NVDA has a class for exactly
  that (`WordDocument_WwN`), so the F7 spelling dialog's "Not in Dictionary" box picked
  up the same markings as the real body. Requiring `_WwG`, the editing surface itself,
  keeps every Word dialog out, whatever window class it turns out to have, because none
  of them can be the document window.
- **window class `_WwG`** — the Word editing surface Outlook composes in. No field on
  the form shares it. Deliberately *not* `RichEdit20W`: NVDA classes every window whose
  class starts with that as a `ContactEditField`, which is what To, Cc and Subject are.
- **`RichEdit20W` with control id 8224** — the one rich edit control that is a body,
  which is how NVDA itself picks out the plain text message.
- a web-view Outlook body, which has none of the above, needs both a name saying it is
  the body and more than one line.

A read-only body, such as the reading pane, is left alone, because the announcement
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

## Documents that take the focus on their own

Dropping document titles is the whole point of the add-on when a page loads. It is
badly wrong at any other moment.

Some pages carry an embedded document — an iframe — that takes the keyboard focus when
the browse cursor comes near it. An "Image Magnify" frame on a product page does it
when you press control+home, or arrow up to the top. **While the focus is inside one of
those, the browser's own control+F does not open the find bar**, and the enter you press
next is taken by browse mode as "activate what is under the cursor", so nothing happens
at all.

NVDA warns you about this by announcing the embedded document as it takes the focus.
The add-on used to drop that announcement along with every other document title in a
browser, which left no way to tell why finding had stopped working — the keyboard had
quietly left the page.

So `_shouldDropObjectSpeech` now drops a document title only when
`_documentAnnouncementIsExpected()` says one is due: while a hooked document call is on
the stack, while the load gate is open, or while a foreground window is being announced.
Those three cover every announcement the add-on set out to remove. A document that names
itself outside all three has taken the focus by itself, and is always spoken.

The residual behaviour is the browser's, not the add-on's: control+F still will not open
the find bar while an embedded document holds the focus. Press escape, or control+home
and then shift+tab out of the frame, and control+F works again. What the add-on owes you
is the announcement that tells you which situation you are in.

## Tracing

Off, and costing nothing, unless a file named `muteBrowseModeTrace.on` exists in your
temp folder. While it does, every keystroke and focus change is appended to
`muteBrowseModeTrace.log` beside it:

```
KEY kb(laptop):enter | script=browseMode…script_activatePosition on ChromeVBuf | focus: name='' role=0 class=Chrome_RenderWidgetHostHWND | ti=ChromeVBuf passThrough=False
```

Each line says which script NVDA resolved for the key, what had the focus, and whether
browse mode was in pass-through. That is the question worth asking about this add-on,
and NVDA's own log cannot answer it with logging turned off, which is the normal
setting. Delete the marker file to stop. Any failure switches tracing off rather than
propagating, so it can never be the thing that breaks a keystroke.

## Control+F opens NVDA's find in a browser

In a **web browser**, control+F opens NVDA's find. NVDA's find searches the browse mode
document from where the cursor is and leaves the cursor on what it found, so the next
line down is the line after the match and NVDA+F3 carries on from there. A browser's
find bar only scrolls the page and puts the keyboard somewhere else entirely. It is also
the one that keeps working: an embedded document holding the focus swallows the
browser's control+F, but NVDA's find never leaves the buffer.

Control+F opens NVDA's find in two kinds of place:

- **a browse mode buffer**, where it searches the whole document, as NVDA's find always
  has;
- **an edit box** — any edit box: a document, a terminal, or anything NVDA marks
  editable — where it searches the field you are in.

Where NVDA has nothing it can search, the key is handed straight back with
`gesture.send()` and does whatever it always did.

### NVDA's find in an edit box

NVDA defines find on `cursorManager.CursorManager`, which only browse mode and a handful
of app modules mix in, so in an ordinary edit box there is nothing for control+F — or for
NVDA+control+F — to invoke at all.

Nothing about the search actually needs a browse mode document, though. `doFindText`
wants somewhere to make a TextInfo, a `find` on that TextInfo, and somewhere to put the
selection, and every editable object already has all three. So `_TextFieldCursorManager`
borrows the object and hands it to NVDA's own `FindDialog` and NVDA's own `doFindText`:
what you get is the find you already know, not a copy of it. `_lastFindText` lives on
`CursorManager` itself, so the text you searched for carries between an edit box, a web
page and NVDA+control+F.

Two details it has to get right. `CursorManager._set_selection` hands `self` to braille
and vision; the adapter hands them the real object instead, because the adapter was made
a moment ago for one search and is about to be dropped. And `TextInfo.find` is
implemented by the offset and UIA text infos but the base class only raises
`NotImplementedError` — so an edit box whose text info cannot search is ruled out
*before* the dialog opens, rather than failing after you have typed what you were
looking for. Password boxes are excluded outright.

### The two check boxes

| Check box | Default | Where control+F opens NVDA's find |
| --- | --- | --- |
| Control+F opens NVDA's find in a web browser | ticked | Web browsers only |
| Bring up NVDA screen reader find when not in Outlook | unticked | Everywhere except Microsoft Outlook |

The second is a widening of the first, so where both are ticked the second wins and
everywhere the first reached is still reached. Untick both and control+F always goes to
the program. NVDA's find stays on NVDA+control+F either way.

**Outlook is the exception in both cases, everywhere and always.** Control+F forwards the
message you are reading, and that is as true of the message body, the subject line and
the address fields as it is of the message list — all of them are edit boxes or browse
mode buffers that would otherwise qualify. `_findSource` tests for Outlook before it
tests for anything else, and neither check box can override it. `_isWebBrowser` rules
Outlook out in its own right as well, because the new Outlook for Windows is a WebView2
application and looks exactly like Chromium from the outside.

### Why the binding comes and goes

`_syncBrowserFindBinding` adds and removes the control+F binding on every foreground
*and* every focus change, rather than binding it once and handing the key back where it
is not wanted. The test it makes is the same one the script makes: is there a browse mode
document here that NVDA's find could search, in a program the check boxes cover. Keying
the binding to that rather than to the program alone is what keeps the key off everything
else — Word, the address bar and the browser's own find bar are never bound, so nothing
is ever injected into them.

A bound key is trapped. `keyboardHandler.internal_keyDownEvent` returns False the moment
`executeGesture` finds a script for it, so the real key down never reaches the program
and `trappedKeys` swallows the key up as well. In a browser that is precisely the point.
Anywhere else it is precisely wrong, because what the program would get instead is a
synthetic keystroke injected from NVDA's main thread an event queue turn later, and a
synthetic one is not the same as a real one:

- `send()` drops any modifier `winUser.getKeyState` reports as already down, so what is
  injected depends on how the key state looks from NVDA's own thread at that moment,
  not on what you are holding;
- a held control+F auto repeats, and every repeat queues another script and another
  injection;
- `send()` keeps `ignoreInjected` up only until the hook thread has seen the injected
  keys or 10 ms have passed, and the key it waits for is `keys[0]`, the first modifier
  rather than the last key up actually sent. Lose that race and NVDA reads its own
  injection back as a real keystroke.

Control+F is Forward in Outlook and has to arrive there as the key you actually pressed.
Between 1.3 and 1.9 the add-on bound control+F system-wide to a script that did nothing
but `gesture.send()`, which bought nothing at all — nothing in NVDA binds plain
control+F, so it already reached every program. That is why it was removed in 2.0, and
why it is back in 2.3 only where it earns its place.

## In Outlook and Chromium browsers

Neither of the two silencing choices stops at the document announcement in Microsoft
Outlook or in Chromium based browsers. There, NVDA also stops speaking:

- **window titles**, so switching to Outlook no longer reads the title of the message
  list window and opening a message no longer reads the title of the message window.
  The one exception is a browser window being switched to, which is read in full — see
  [The window title itself](#the-window-title-itself) above;
- **dialogs**, so opening an Outlook message no longer says the word "dialog". A
  message opens inside a dialog, and NVDA announces it while walking down the focus
  ancestors, which happens before the message document exists;
- **document titles**, so a browser tab no longer announces the page title — but only
  while a page is loading, while a document is being entered, or while a window you
  have just switched to is being announced. See
  [Documents that take the focus on their own](#documents-that-take-the-focus-on-their-own)
  below;
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

That writes `muteBrowseMode-2.7.nvda-addon` next to `build.py`. Open it to install,
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
