# Mute Browse Mode

An NVDA add-on that stops NVDA reading out the name of a web page or an email every
time one opens.

NVDA loves to talk when a page opens. It says the name of the page, then the word
"document", then the first line. Every single time. This add-on tells it to be quiet,
so you can start reading the page yourself with the arrow keys.

It also tidies up some noisy bits of Microsoft Outlook, and it lets you use NVDA's own
find box with control+F.

Needs NVDA 2026.1 or later.

## Get it

Download the `.nvda-addon` file from the [Releases](../../releases/latest) page and
open it. Then restart NVDA.

## Where the settings are

Press NVDA+n for the NVDA menu, then Preferences, then Settings, then Browse Mode.
Everything this add-on adds is on that page, underneath NVDA's own settings.

The first setting starts on Normal, so NVDA carries on as usual until you change it.
The page summary is the one thing that is switched on from the start.

## What you get

**Quiet when a page opens.** No page name, no first line, no "Loading document". You can
have three quick beeps instead, so you still know when the page is ready, or leave NVDA
exactly as it was.

**A page summary instead.** When a web page finishes loading, NVDA can tell you how big
it is rather than what it is called: "Loading complete. Page has 8 regions, 57 headings
and 196 links." Outlook never gets one, because Outlook is where this add-on is meant to
be quiet.

**Swapping windows is never cut off.** NVDA always finishes telling you where you have
landed before the quiet starts again, however fast you flick through your windows.

**Landing in Outlook** answers with a short description of where you are — "Inbox, list,
Josh Kennedy build 1.4, list item, 3 of 57" — instead of the window title.

**The list of emails** no longer says the word "Subject" in front of every single
subject.

**Writing an email** says "You are now in the message body, type a message." when you tab
into the box you type in, instead of reading out the empty page behind it.

**The F7 spelling checker** says "Not in Dictionary", then the word it has stopped on,
then spells it out a little slower than usual. When it is done you hear "Spell check is
complete. OK button" and nothing else.

**Links on their own line**, if you tick that box, so the down arrow reaches every link
and button in an Outlook email one at a time.

**Control+F opens NVDA's find**, if you tick that box. It is nicer for reading with than
a browser's find bar: it starts from your cursor and leaves your cursor on what it found.
One box covers web browsers, the other covers everywhere except Outlook — where control+F
forwards an email, and always will.

The full documentation is in the add-on itself. Open the NVDA menu, then Tools, then
Manage Add-ons, pick Mute Browse Mode and press the Add-on Help button.

## Building it yourself

The add-on is the contents of the `addon` folder, zipped up. There is a script for it:

```
python build.py
```

It writes `muteBrowseMode-<version>.nvda-addon` beside itself. Any Python 3 will do; it
does not need NVDA, SCons or the add-on template.

## License

GNU General Public License, version 2. See [LICENSE](LICENSE).
