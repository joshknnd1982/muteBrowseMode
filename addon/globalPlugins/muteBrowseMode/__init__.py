# -*- coding: utf-8 -*-
# Mute Browse Mode, an NVDA add-on.
# Copyright (C) 2026 Josh Kennedy
# This file is covered by the GNU General Public License, version 2.

"""Stops NVDA announcing a browse mode document every time one loads or is entered.

NVDA speaks a document at several distinct moments, and they are not all part of a
document load:

* a virtual buffer starting and finishing loading ("Loading document...", "Refreshed"),
* a browse mode document being entered for the first time (the name, the word
  "document", and the first line of the buffer),
* focus arriving in a browse mode document from outside it, which NVDA answers by
  reading the line the focus landed on.

That last one is what reads the first line of an Outlook message, and because it is a
focus event rather than a load, hooking the load is not enough to stop it.

In Outlook and in Chromium based browsers the add-on goes further and never speaks
window, dialog or document titles, toasts, or live region "flash" messages, because in
those two applications they are noise rather than information. The dialog is part of
this: an Outlook message opens inside one, and NVDA names it, and says the word
"dialog", on its way down to the focus. That happens before the document exists, so it
is dropped by role rather than gated.

In a web browser, and only when Outlook is *not* the program the user is in, the
silence is replaced by a page summary once the buffer has finished loading: "Loading
complete", then how many regions, headings and links the page has. Outlook is the one
place that must stay silent, so every page summary is gated on a check of the current
application. That check is silent: its answer only ever reaches NVDA's log.

The gate is deadline based rather than a counter, so a bug or an exception can never
leave NVDA permanently mute: the worst case is a few seconds of silence that expires
on its own. Any input gesture closes the gate immediately, so pressing a key always
gets normal speech back straight away. Braille is deliberately untouched: everything
here suppresses speech only, and the calls that would have spoken still run, so
braille, vision and NVDA's own caches carry on as usual.
"""

import ctypes
import ctypes.wintypes
import os
import re
import tempfile
import time

import addonHandler
import api
import browseMode
import config
import controlTypes
import core
import cursorManager
import globalPluginHandler
import inputCore
import scriptHandler
import speech
import speech.speech
import textInfos
import tones
import ui
import virtualBuffers
import winUser
import wx
from gui import guiHelper, settingsDialogs
from logHandler import log

try:
	addonHandler.initTranslation()
except Exception:
	# Running outside an installed add-on (e.g. from source). NVDA installs a
	# gettext ``_`` into builtins anyway, so this is only belt and braces.
	log.debugWarning("Mute Browse Mode: translations unavailable", exc_info=True)

try:
	from speech.commands import CallbackCommand as _CallbackCommand
except Exception:
	# Without it the foreground announcement cannot be told exactly when it has been
	# spoken, and falls back to running to its ceiling. Everything else still works.
	_CallbackCommand = None
	log.debugWarning("Mute Browse Mode: CallbackCommand unavailable", exc_info=True)

try:
	from speech.commands import EndUtteranceCommand as _EndUtteranceCommand
	from speech.commands import RateCommand as _RateCommand
except Exception:
	# Without these a misspelled word is still spelled out, just not more slowly.
	_EndUtteranceCommand = None
	_RateCommand = None
	log.debugWarning("Mute Browse Mode: speech commands unavailable", exc_info=True)

if "ngettext" not in globals():
	# initTranslation puts ngettext in this module's globals. NVDA only installs ``_``
	# into builtins, so if that call failed there would otherwise be no plural form.
	def ngettext(singular, plural, n):
		return singular if n == 1 else plural


#: Section this add-on owns in nvda.ini.
CONF_SECTION = "muteBrowseMode"

MODE_SILENCE = "silence"
MODE_TONES = "tones"
MODE_NORMAL = "normal"

#: Order matters: this is the order of the entries in the combo box.
MODES = (MODE_SILENCE, MODE_TONES, MODE_NORMAL)

SUMMARY_SPEAK = "speak"
SUMMARY_TONES = "tones"
SUMMARY_NORMAL = "normal"

#: Order matters: this is the order of the entries in the combo box.
SUMMARY_MODES = (SUMMARY_SPEAK, SUMMARY_TONES, SUMMARY_NORMAL)

config.conf.spec[CONF_SECTION] = {
	"mode": 'option("silence", "tones", "normal", default="normal")',
	"pageSummary": 'option("speak", "tones", "normal", default="speak")',
	"announceLoadingComplete": "boolean(default=True)",
	"linksOnOwnLine": "boolean(default=False)",
	"browserFind": "boolean(default=True)",
	"findWhenNotOutlook": "boolean(default=False)",
}


def getModeLabels():
	"""Combo box entries, in the same order as L{MODES}."""
	return [
		# Translators: A choice in the "Mute browse mode" combo box in Browse Mode settings.
		_("Silence all browsing"),
		# Translators: A choice in the "Mute browse mode" combo box in Browse Mode settings.
		_("Play tones"),
		# Translators: A choice in the "Mute browse mode" combo box in Browse Mode settings.
		_("Normal"),
	]


def getSummaryModeLabels():
	"""Combo box entries, in the same order as L{SUMMARY_MODES}."""
	return [
		# Translators: A choice in the "page summary" combo box in Browse Mode settings.
		_("Speak the summary"),
		# Translators: A choice in the "page summary" combo box in Browse Mode settings.
		_("Play tones"),
		# Translators: A choice in the "page summary" combo box in Browse Mode settings.
		_("Normal"),
	]


def _getOption(key, valid, default):
	"""One of this add-on's options, falling back if the config is missing or odd."""
	try:
		value = config.conf[CONF_SECTION][key]
	except Exception:
		return default
	return value if value in valid else default


def getMode():
	return _getOption("mode", MODES, MODE_NORMAL)


def setMode(mode):
	config.conf[CONF_SECTION]["mode"] = mode


def getSummaryMode():
	return _getOption("pageSummary", SUMMARY_MODES, SUMMARY_SPEAK)


def setSummaryMode(mode):
	config.conf[CONF_SECTION]["pageSummary"] = mode


def getAnnounceLoadingComplete():
	try:
		return bool(config.conf[CONF_SECTION]["announceLoadingComplete"])
	except Exception:
		return True


def setAnnounceLoadingComplete(enabled):
	config.conf[CONF_SECTION]["announceLoadingComplete"] = bool(enabled)


def getLinksOnOwnLine():
	try:
		return bool(config.conf[CONF_SECTION]["linksOnOwnLine"])
	except Exception:
		return False


def setLinksOnOwnLine(enabled):
	config.conf[CONF_SECTION]["linksOnOwnLine"] = bool(enabled)


def getBrowserFind():
	try:
		return bool(config.conf[CONF_SECTION]["browserFind"])
	except Exception:
		return True


def setBrowserFind(enabled):
	config.conf[CONF_SECTION]["browserFind"] = bool(enabled)


def getFindWhenNotOutlook():
	try:
		return bool(config.conf[CONF_SECTION]["findWhenNotOutlook"])
	except Exception:
		return False


def setFindWhenNotOutlook(enabled):
	config.conf[CONF_SECTION]["findWhenNotOutlook"] = bool(enabled)


### The applications we treat specially

TARGET_OUTLOOK = "outlook"
TARGET_CHROMIUM = "chromium"

#: Executable names, lower case and without the extension, that count as Outlook.
_OUTLOOK_APP_NAMES = frozenset((
	"outlook",  # classic desktop Outlook
	"olk",  # the new Outlook for Windows
	"hxoutlook",  # Outlook / Mail from the Microsoft Store
	"hxmail",
	"msoutlook",
))

#: Window classes Outlook gives its own top level windows: one class for the classic
#: desktop Outlook's main window and for every message opened in a window of its own,
#: and the host window the new Outlook for Windows puts its interface inside.
#:
#: These are here because the executable behind the object with the focus is not always
#: Outlook even when Outlook is the program the user is in. Outlook renders more and
#: more of itself in an embedded Edge web view — in the newest builds the message body
#: as well — and every window a web view makes belongs to msedgewebview2.exe, which is a
#: different process from Outlook's own and is Chromium by every test there is. The top
#: level window is the one thing that is still Outlook's, whatever is embedded in it.
_OUTLOOK_WINDOW_CLASSES = frozenset((
	"rctrl_renwnd32",  # classic Outlook: the main window and every message window
	"Outlook Host",  # the new Outlook for Windows
))

#: Executable names of Chromium based browsers. The window class check below catches
#: nearly all of them on its own; this is for objects with no usable window class.
_CHROMIUM_APP_NAMES = frozenset((
	"chrome",
	"chromium",
	"chromium-browser",
	"chrome_proxy",
	"msedge",
	"msedgewebview2",
	"microsoftedge",
	"microsoftedgecp",
	"brave",
	"opera",
	"vivaldi",
	"browser",  # Yandex Browser ships as browser.exe
	"whale",
	"iron",
	"thorium",
	"slimjet",
	"centbrowser",
	"epic",
	"dragon",
	"comodo_dragon",
	"maxthon",
	"ungoogled-chromium",
	"arc",
))

#: Window classes every Chromium build uses, for its top level windows and for the
#: window the page itself is rendered into. Matching on these rather than on a list of
#: executables means new browsers and forks are covered without an add-on update, and
#: it is also how the new Outlook is recognised, since that is a WebView2 application.
#: Electron applications are covered by the same classes.
_CHROMIUM_WINDOW_CLASSES = frozenset((
	"Chrome_WidgetWin_0",
	"Chrome_WidgetWin_1",
	"Chrome_RenderWidgetHostHWND",
))

#: Executable names of Gecko based browsers. Firefox is not Chromium, so it is not one
#: of the applications the add-on is aggressive in, but a page loading in it is still a
#: page, and it still gets a page summary.
_GECKO_APP_NAMES = frozenset((
	"firefox",
	"firefox-esr",
	"waterfox",
	"librewolf",
	"floorp",
	"zen",
	"palemoon",
	"basilisk",
	"seamonkey",
	"icecat",
	"mercury",
))

#: Window classes every Gecko build uses.
_GECKO_WINDOW_CLASSES = frozenset((
	"MozillaWindowClass",
	"MozillaDialogClass",
	"MozillaContentWindowClass",
))


def _appNameOf(obj):
	try:
		return (obj.appModule.appName or "").lower()
	except Exception:
		return ""


def _windowClassOf(obj):
	try:
		return getattr(obj, "windowClassName", "") or ""
	except Exception:
		return ""


def _isOutlook(obj):
	"""Whether C{obj} itself belongs to any version of Microsoft Outlook."""
	return obj is not None and _appNameOf(obj) in _OUTLOOK_APP_NAMES


def _rootWindowOf(obj):
	"""The handle of the top level window C{obj} sits in, or 0 for none."""
	try:
		hwnd = obj.windowHandle
	except Exception:
		return 0
	if not hwnd:
		return 0
	try:
		import winUser

		return winUser.getAncestor(hwnd, getattr(winUser, "GA_ROOT", 2)) or 0
	except Exception:
		return 0


def _processIDOf(obj):
	"""The process C{obj} belongs to, or C{None} if it will not say."""
	try:
		return obj.processID
	except Exception:
		return None


def _appNameOfProcess(processID):
	"""The application a process is, named the way L{_appNameOf} names one.

	Asked of NVDA's own table of running applications, which is keyed by process and
	already holds an entry for every program NVDA has looked at. So this is a dictionary
	lookup for the window in front, and only walks the process list — which is what makes
	the question expensive — for a process NVDA has never seen before.
	"""
	if not processID:
		return ""
	try:
		import appModuleHandler

		return (appModuleHandler.getAppModuleFromProcessID(processID).appName or "").lower()
	except Exception:
		return ""


def _isInOutlookWindow(obj):
	"""Whether C{obj} is anywhere inside a window belonging to Microsoft Outlook.

	L{_isOutlook} asks what application the object belongs to, and that is no longer the
	same question. Outlook renders parts of itself in an embedded Edge web view, and the
	newest builds render the message body in one too: everything inside it belongs to
	msedgewebview2.exe, names that as its application, and is indistinguishable from a
	Chromium browser by every test the add-on makes. That is what took control+F off
	Outlook and opened NVDA's find in the middle of a message.

	The window the user is looking at is the answer to that. Whatever is embedded inside
	it, the top level window is Outlook's own, and it says so twice over: by its class,
	and by the process that owns it.

	Runs on every focus change, so it stops as soon as it has an answer. Where the top
	level window belongs to the same program as the object itself — which is every
	ordinary window, embedded or not — the process is not looked up at all: that question
	was answered on the first line.
	"""
	if _isOutlook(obj):
		return True
	root = _rootWindowOf(obj)
	if not root:
		return False
	try:
		import winUser

		if winUser.getClassName(root) in _OUTLOOK_WINDOW_CLASSES:
			return True
		processID = winUser.getWindowThreadProcessID(root)[0]
	except Exception:
		log.debugWarning("Mute Browse Mode: could not read the top level window", exc_info=True)
		return False
	if not processID or processID == _processIDOf(obj):
		return False
	return _appNameOfProcess(processID) in _OUTLOOK_APP_NAMES


def outlookIsCurrent(obj=None):
	"""Whether Microsoft Outlook is the program the user is currently in.

	Answered from the foreground window and the focus, so it is the running program
	that decides, not the object that happened to raise an event. C{obj}, when given,
	is checked first: a document belonging to Outlook counts as Outlook even in the
	moment before its window has become the foreground one.

	Each of the three is judged by the window it is in rather than by the process it
	belongs to, because an Outlook message body may be an embedded web view and so belong
	to another process altogether. See L{_isInOutlookWindow}.

	Nothing here reaches the synthesiser, so the check itself is never heard.
	"""
	if obj is not None and _isInOutlookWindow(obj):
		return True
	for getter in (api.getForegroundObject, api.getFocusObject):
		try:
			candidate = getter()
		except Exception:
			continue
		if _isInOutlookWindow(candidate):
			return True
	return False


def _targetOf(obj):
	"""Which of the two applications C{obj} belongs to, or C{None} for anything else.

	Outlook is asked first and by the window, because an Outlook message rendered in an
	embedded web view answers every Chromium test there is — see L{_isInOutlookWindow}.
	The two are not interchangeable here: a browser window being switched to is read out
	in full, and an Outlook one is not.
	"""
	if obj is None:
		return None
	if _isInOutlookWindow(obj):
		return TARGET_OUTLOOK
	if _appNameOf(obj) in _CHROMIUM_APP_NAMES:
		return TARGET_CHROMIUM
	if _windowClassOf(obj) in _CHROMIUM_WINDOW_CLASSES:
		return TARGET_CHROMIUM
	return None


def _isWebBrowser(obj):
	"""Whether C{obj} belongs to a browser or an Electron application, and not Outlook.

	Outlook is excluded before anything else, and by the window it is in: the new Outlook
	for Windows is a WebView2 application, and the classic one now renders the message
	itself in a web view, so both look exactly like Chromium from the outside.
	"""
	if obj is None:
		return False
	target = _targetOf(obj)
	if target is not None:
		# Chromium, or an Outlook that is only wearing Chromium's clothes.
		return target == TARGET_CHROMIUM
	if _appNameOf(obj) in _GECKO_APP_NAMES:
		return True
	return _windowClassOf(obj) in _GECKO_WINDOW_CLASSES


### The column headings the Outlook message list reads out
#
# A row of the Outlook message list is announced one column at a time, and each column
# is read out with its heading in front of it: "unread, From MacDailyNews, Subject Amy
# Sedaris joins Ben Stiller". The subject is the one column that never needs saying: it
# is the line of the message the user came to hear, it is by far the longest thing on
# the row, and the heading in front of it is repeated on every single message.
#
# Only a heading standing in front of something is dropped, which is what keeps the
# subject *field* on a message being written: NVDA speaks that as its own item in the
# sequence, "Subject:" or "Subject" with the value following separately, and neither is
# a heading with the value run on behind it.

#: Column headings dropped from the front of a cell, lower case.
_SUPPRESSED_COLUMN_LABELS = ("subject",)

#: A heading at the start of an item or straight after a comma, with a value behind it.
_COLUMN_LABEL_RE = re.compile(
	r"(^|[,;]\s*)(?:%s)\s+(?=\S)" % "|".join(_SUPPRESSED_COLUMN_LABELS),
	re.IGNORECASE,
)


def _isOutlookRowFocus():
	"""Whether the focus is on a row of one of Outlook's lists.

	What keeps the heading filter off everything else. A line of message text that
	happens to begin "subject to change" is read with the focus on the document, and the
	subject field of a message being written is an edit box whose name is the bare word
	with the value spoken separately after it, so neither of them is touched. Only the
	thing that actually reads out a heading in front of every column is.

	Asked two ways, because the message list is drawn differently in different versions
	of Outlook and the role a row comes out as is not the same in all of them: the role
	if it is one a row has, and otherwise whether the focus is carrying a heading and a
	value run together in its own name, which is exactly what a row of columns is and
	what nothing else in Outlook is.
	"""
	try:
		focus = api.getFocusObject()
	except Exception:
		return False
	if not _isInOutlookWindow(focus):
		return False
	if getattr(focus, "role", None) in _ROW_ROLES:
		return True
	try:
		return bool(_COLUMN_LABEL_RE.search(focus.name or ""))
	except Exception:
		return False


def _filterColumnLabels(args, kwargs):
	"""Take the suppressed column headings out of a speech sequence, in Outlook only.

	Where the focus is gets asked only once something has actually matched, so the
	ordinary case — everywhere else, and every other utterance in Outlook — costs one
	failed regular expression search per string and nothing else. Anything in the
	sequence that is not a string is a speech command and is passed through untouched.
	"""
	if args:
		sequence = args[0]
	elif "speechSequence" in kwargs:
		sequence = kwargs["speechSequence"]
	else:
		return args, kwargs
	if not isinstance(sequence, list) or not sequence:
		return args, kwargs
	if not any(isinstance(item, str) and _COLUMN_LABEL_RE.search(item) for item in sequence):
		return args, kwargs
	if not _isOutlookRowFocus():
		return args, kwargs
	filtered = [
		_COLUMN_LABEL_RE.sub(r"\1", item) if isinstance(item, str) else item
		for item in sequence
	]
	if args:
		return (filtered,) + tuple(args[1:]), kwargs
	kwargs = dict(kwargs)
	kwargs["speechSequence"] = filtered
	return args, kwargs


### Roles and reasons

def _members(enum, names):
	"""The named members of C{enum} that this NVDA actually has."""
	return frozenset(m for m in (getattr(enum, name, None) for name in names) if m is not None)


#: Roles that are really just a container: a window, a dialog or a document, whose name
#: is its title. NVDA speaks these when a window comes to the foreground and as it walks
#: down the focus ancestors, which is where the Outlook message list window title, the
#: message window title, and the bare word "dialog" an opening message announces all
#: come from.
#:
#: DIALOG matters as much as WINDOW here. An Outlook message opens inside a dialog, and
#: NVDA announces it from ``NVDAObject.event_focusEntered`` while walking down to the
#: focus. That walk finishes before ``event_treeInterceptor_gainFocus`` runs, so it
#: happens before any of the browse mode hooks below have opened the gate, and dropping
#: the announcement here is the only thing that catches it.
_TITLE_ROLES = _members(
	controlTypes.Role,
	(
		"WINDOW",
		"PANE",
		"FRAME",
		"INTERNALFRAME",
		"DIALOG",
		"DOCUMENT",
		"APPLICATION",
		"PROPERTYPAGE",
	),
)

#: The subset of the above whose name is the title of a whole window, as opposed to a
#: dialog or a document. These are the ones that come back when a window is switched
#: to, because that title is the only thing telling the user where they have landed.
#: DIALOG and DOCUMENT stay out of it: the dialog is the word an opening Outlook
#: message says, and the document is the page title the add-on exists to silence.
_WINDOW_TITLE_ROLES = _members(
	controlTypes.Role,
	("WINDOW", "PANE", "FRAME", "INTERNALFRAME", "APPLICATION"),
)

#: The subset that names a document rather than a window. NVDA announces one of these
#: when a page finishes loading and when a browser window is switched to, which is the
#: noise this add-on exists to remove — but it announces the very same thing when the
#: focus moves into an *embedded* document part way through a session, and that is not
#: noise at all. It means the keyboard has gone somewhere the user did not put it, and
#: while it is there the browser's own shortcuts may not reach the browser: an iframe
#: that has taken the focus swallows control+F, so the find bar never opens and the
#: enter that follows goes to browse mode instead. Silencing that left the user with no
#: way to tell why finding had stopped working. See L{_documentAnnouncementIsExpected}.
_DOCUMENT_TITLE_ROLES = _members(controlTypes.Role, ("DOCUMENT",))

#: Roles used for toasts, flash messages and other transient announcements.
_ALERT_ROLES = _members(controlTypes.Role, ("ALERT", "TOOLTIP", "HELPBALLOON", "NOTIFICATION"))

#: Roles a row of a list or a table has. The Outlook message list uses one of these
#: whichever way it is being drawn, and it is the only thing whose announcement carries a
#: heading in front of every column. See L{_isOutlookRowFocus}.
_ROW_ROLES = _members(
	controlTypes.Role,
	("DATAITEM", "LISTITEM", "TREEVIEWITEM", "TABLEROW", "TABLECELL"),
)

#: Reasons that mean "NVDA decided to say this". Everything else is left alone: QUERY is
#: the user asking with NVDA+tab, SAYALL and MESSAGE are explicit, and ONLYCACHE must
#: never be dropped because browse mode uses it to keep its property cache honest.
_AUTOMATIC_REASONS = _members(
	controlTypes.OutputReason,
	("FOCUS", "FOCUSENTERED", "CHANGE", "CARET"),
)


### The speech gate

#: Ceiling while a hooked call is still on the stack.
_IN_CALL_GATE = 6.0
#: How long the gate is held after a hooked call returns. NVDA queues part of the
#: document announcement onto the main queue, so it does not all happen inline.
_TRAILING_GATE = 1.5
#: Ceiling while a virtual buffer is loading. This is what silences "Loading
#: document...", which NVDA speaks from a timer part way through the load.
_LOAD_GATE = 15.0
#: Longest a single silenced call may hold speech shut, whatever happens inside it.
_MUTE_CEILING = 5.0

#: How long a foreground announcement may hold the deadline gate off, whatever
#: happens. A ceiling, not a target: the announcement normally ends itself the moment
#: the synthesiser finishes saying it.
_ANNOUNCE_CEILING = 8.0
#: How long after a foreground change we wait to see whether NVDA is going to say
#: anything at all before giving up on the announcement, in milliseconds.
_ANNOUNCE_SETTLE = 300
#: How long after the last utterance has been spoken the announcement stays open, in
#: milliseconds, in case NVDA has more of it still to come.
_ANNOUNCE_TAIL = 250

#: Monotonic timestamp at which speech is allowed through again. 0 means open.
_gateUntil = 0.0
#: Depth of nested "we are inside a document announcement right now" wrappers, and the
#: latest moment they may keep speech shut. This is the gate that actually silences the
#: document: NVDA speaks all of it inline, so being inside the call is enough, and it
#: cannot bleed into whatever NVDA says next.
_inCallDepth = 0
_inCallUntil = 0.0
#: Depth of nested "silence just this call" wrappers, and the latest moment any of them
#: may keep speech shut. Both have to agree, so even a lost decrement expires by itself.
_muteDepth = 0
_muteUntil = 0.0
#: Depth of nested "this is the add-on speaking" wrappers. Speech from inside one of
#: these is never gated: it is the add-on's own announcement, not NVDA's.
_bypassDepth = 0
#: Latest moment the foreground announcement may still be running, and how many
#: utterances of it the synthesiser has not reached yet. While this is live the
#: deadline gate stands down, so alt+tab is never cut off.
_announceUntil = 0.0
_announcePending = 0
#: Which announcement those two belong to. Bumped on every window switch, so that a
#: callback for an utterance of the window just left cannot be counted against the
#: window just arrived in.
_announceGeneration = 0


def _openGate(seconds):
	global _gateUntil
	_gateUntil = time.monotonic() + seconds


def _closeGate():
	global _gateUntil
	_gateUntil = 0.0


def _enterDocumentCall():
	global _inCallDepth, _inCallUntil
	_inCallDepth += 1
	_inCallUntil = max(_inCallUntil, time.monotonic() + _IN_CALL_GATE)


def _exitDocumentCall():
	global _inCallDepth, _inCallUntil
	_inCallDepth -= 1
	if _inCallDepth <= 0:
		_inCallDepth = 0
		_inCallUntil = 0.0


class _hardMute:
	"""Context manager dropping speech for the duration of one call.

	Used where the call has to happen for braille and NVDA's own bookkeeping, but must
	not reach the synthesiser. Unlike the deadline gate this is not cleared by a
	gesture, because it only ever spans a single synchronous call.
	"""

	def __enter__(self):
		global _muteDepth, _muteUntil
		_muteDepth += 1
		_muteUntil = max(_muteUntil, time.monotonic() + _MUTE_CEILING)
		return self

	def __exit__(self, *exc):
		global _muteDepth, _muteUntil
		_muteDepth -= 1
		if _muteDepth <= 0:
			_muteDepth = 0
			_muteUntil = 0.0
		return False


class _ownSpeech:
	"""Context manager letting the add-on's own announcements past the gate.

	The page summary is spoken in the moment the gate is holding NVDA's document
	announcement shut, which is exactly the point: it is what replaces it.
	"""

	def __enter__(self):
		global _bypassDepth
		_bypassDepth += 1
		return self

	def __exit__(self, *exc):
		global _bypassDepth
		_bypassDepth = max(0, _bypassDepth - 1)
		return False


def _resetGates():
	global _muteDepth, _muteUntil, _bypassDepth, _inCallDepth, _inCallUntil
	_muteDepth = 0
	_muteUntil = 0.0
	_bypassDepth = 0
	_inCallDepth = 0
	_inCallUntil = 0.0
	_endAnnouncement()
	_closeGate()


def _sayAllRunning():
	try:
		from speech.sayAll import SayAllHandler

		return SayAllHandler.isRunning()
	except Exception:
		return False


def _isGated():
	if _bypassDepth > 0:
		return False
	now = time.monotonic()
	if _muteDepth > 0 and now < _muteUntil:
		return True
	inCall = _inCallDepth > 0 and now < _inCallUntil
	# The deadline gate stands down while NVDA is announcing a new foreground window.
	# Only the in-call gate above survives that, and it only covers the document
	# announcement itself, which NVDA speaks entirely inline.
	deadline = _gateUntil > 0.0 and now < _gateUntil and not _announcingNow()
	if not inCall and not deadline:
		return False
	if _sayAllRunning():
		# Say all is an explicit "read this to me" request, including read on page
		# load. Never swallow it, and close the gate so its start is not clipped.
		_closeGate()
		return False
	return True


### The foreground announcement
#
# Alt+tab makes NVDA announce, in this order, the title of the window being switched
# to, then the browse mode document if there is one, then the control the focus landed
# on. Only the middle one is ours to silence, and NVDA speaks all of it inline, so the
# in-call gate above is enough for it. The deadline gate is not: a page that is still
# loading holds it open for fifteen seconds, and the trailing gate after a document
# announcement holds it for another second and a half, and either of those would eat
# the announcement of the control the user has just switched to.
#
# So a foreground change stands the deadline gate down until NVDA has finished
# speaking. "Finished" is the synthesiser's word, not a guess: a CallbackCommand is
# appended to each utterance we let through, and NVDA runs it when speech actually
# reaches that point.
#
# Each announcement is numbered, and every callback carries the number of the
# announcement it was made for. It has to: NVDA hands those callbacks back through the
# event queue (``_onSynthIndexReached`` queues ``_handleIndex``), so one belonging to
# the window the user has just left can easily arrive after the next switch has begun.
# An unnumbered callback would then count against the new announcement and end it
# early, which is exactly how the title of the window being switched to ended up only
# half read.


def _announcingNow():
	return _announceUntil > 0.0 and time.monotonic() < _announceUntil


def _beginAnnouncement():
	"""A new window has come to the front, so stop silencing until it has been read."""
	global _announceUntil, _announcePending, _announceGeneration
	_announceGeneration += 1
	_announcePending = 0
	_announceUntil = time.monotonic() + _ANNOUNCE_CEILING
	core.callLater(_ANNOUNCE_SETTLE, _endAnnouncementIfIdle, _announceGeneration)


def _endAnnouncement():
	global _announceUntil, _announcePending, _announceGeneration
	_announceUntil = 0.0
	_announcePending = 0
	# Anything still in flight belonged to the announcement that has just ended.
	_announceGeneration += 1


def _endAnnouncementIfIdle(generation):
	"""End the announcement, unless it is still being spoken or has been replaced."""
	if generation == _announceGeneration and _announcePending <= 0:
		_endAnnouncement()


def _announcementSpoken(generation):
	"""The synthesiser has reached the end of one utterance of the announcement.

	Waits a moment before ending, rather than ending on the spot, so that an
	announcement NVDA is still adding to is not closed off in a gap between two of its
	own utterances.
	"""
	global _announcePending
	if generation != _announceGeneration:
		# Left over from a window the user has already switched away from.
		return
	_announcePending -= 1
	if _announcePending <= 0:
		core.callLater(_ANNOUNCE_TAIL, _endAnnouncementIfIdle, generation)


def _tagAnnouncement(args, kwargs):
	"""Append the "this has now been spoken" callback to a speech sequence.

	Returns the arguments to call the real ``speak`` with. If anything about the call
	is not what we expect, they come back untouched: a missed tag only means the
	announcement runs to its ceiling instead of ending exactly on time.
	"""
	global _announcePending
	if _CallbackCommand is None:
		return args, kwargs
	if args:
		sequence = args[0]
	elif "speechSequence" in kwargs:
		sequence = kwargs["speechSequence"]
	else:
		return args, kwargs
	if not isinstance(sequence, list) or not sequence:
		return args, kwargs
	generation = _announceGeneration
	tagged = sequence + [
		_CallbackCommand(
			lambda: _announcementSpoken(generation),
			name="muteBrowseMode.announced",
		),
	]
	_announcePending += 1
	if args:
		return (tagged,) + tuple(args[1:]), kwargs
	kwargs = dict(kwargs)
	kwargs["speechSequence"] = tagged
	return args, kwargs


### The ready chime

#: (hz, milliseconds) triples, ascending, played when a document is ready.
_TONE_SEQUENCE = ((440, 55), (587, 55), (784, 70))
#: Gap between tones, on top of each tone's own length.
_TONE_GAP = 15
#: Opening a document fires several of our hooks. Only chime once per document.
_TONE_DEBOUNCE = 1.0

_lastTones = 0.0


def _playReadyTones():
	"""Three quick ascending tones, scheduled so they play in sequence."""
	global _lastTones
	now = time.monotonic()
	if now - _lastTones < _TONE_DEBOUNCE:
		return
	_lastTones = now
	delay = 0
	for hz, ms in _TONE_SEQUENCE:
		core.callLater(delay, tones.beep, hz, ms)
		delay += ms + _TONE_GAP


### The page summary

#: Attribute set on a virtual buffer that has started loading and has not yet had its
#: summary announced. Kept on the buffer rather than in a module level variable because
#: several tabs and frames load at once and only the focused one is announced.
_ARMED_ATTR = "_muteBrowseModeSummaryArmed"

#: Node types counted, in the order they are announced.
_SUMMARY_TYPES = ("landmark", "heading", "link")

#: How long after the load finishes the summary is announced. NVDA queues the tail of
#: its own document announcement onto the main queue, so this waits for that to have
#: come and gone rather than talking over it.
_SUMMARY_DELAY = 400
#: Never count more than this many of any one element.
_SUMMARY_MAX = 1500
#: Wall clock budget for counting a whole page. Each element found is a separate call
#: into the virtual buffer, so a pathological page must not be allowed to stall NVDA.
_SUMMARY_BUDGET = 1.5
#: How often, in elements, the budget is checked.
_SUMMARY_BUDGET_EVERY = 50

#: The pending wx.CallLater, so a second load can replace the first one's summary.
_summaryTimer = None


def _cancelSummary():
	global _summaryTimer
	timer, _summaryTimer = _summaryTimer, None
	if timer is None:
		return
	try:
		timer.Stop()
	except Exception:
		pass


def _armSummary(buffer, args=None, kwargs=None):
	"""A buffer has started loading, so it is owed one summary."""
	try:
		setattr(buffer, _ARMED_ATTR, True)
	except Exception:
		log.debugWarning("Mute Browse Mode: could not arm the page summary", exc_info=True)


def _isFocusedBuffer(buffer):
	"""Whether C{buffer} is the document the user is actually in.

	The same check NVDA makes before reporting a document it has just loaded. Several
	tabs load at once, and only the one being looked at has anything to say.
	"""
	try:
		return api.getFocusObject().treeInterceptor is buffer
	except Exception:
		return False


def _scheduleSummary(buffer, args=(), kwargs=None):
	"""A buffer has finished loading, so queue its summary."""
	global _summaryTimer
	if getSummaryMode() == SUMMARY_NORMAL or not getattr(buffer, _ARMED_ATTR, False):
		return
	if not _isFocusedBuffer(buffer):
		# A background tab. Checked here as well as when the summary comes round, so
		# that one cannot take the pending slot away from the tab being looked at.
		return
	_cancelSummary()
	_summaryTimer = core.callLater(_SUMMARY_DELAY, _announceSummary, buffer)


def _bufferTextLength(buffer):
	"""How many characters C{buffer} holds, or C{None} if that cannot be told.

	A virtual buffer that finished loading empty is one NVDA is still waiting on: it
	handles that by reporting the document again from ``event_documentLoadComplete``.
	Counting an empty buffer would announce a page of nothing.
	"""
	try:
		import NVDAHelper

		handle = buffer.VBufHandle
		if not handle:
			return 0
		return int(NVDAHelper.localLib.VBuf_getTextLength(handle))
	except Exception:
		return None


def _countElements(buffer):
	"""How many of each of L{_SUMMARY_TYPES} the buffer holds.

	Returns a C{{nodeType: (count, capped)}} mapping, where C{capped} means the count
	stopped early and is a floor rather than a total.
	"""
	deadline = time.monotonic() + _SUMMARY_BUDGET
	counts = {}
	for nodeType in _SUMMARY_TYPES:
		count = 0
		capped = False
		try:
			# Raised eagerly by NVDA when a backend cannot search for this node type,
			# so it has to be caught around the call and not just around the loop.
			nodes = buffer._iterNodesByType(nodeType)
			for _node in nodes:
				count += 1
				if count >= _SUMMARY_MAX:
					capped = True
					break
				if count % _SUMMARY_BUDGET_EVERY == 0 and time.monotonic() > deadline:
					capped = True
					break
		except NotImplementedError:
			pass
		except Exception:
			log.debugWarning("Mute Browse Mode: could not count %s" % nodeType, exc_info=True)
		counts[nodeType] = (count, capped)
	return counts


def _countPhrase(nodeType, count, capped):
	if nodeType == "landmark":
		# Translators: Part of the page summary announced when a page has loaded.
		text = ngettext("%d region", "%d regions", count) % count
	elif nodeType == "heading":
		# Translators: Part of the page summary announced when a page has loaded.
		text = ngettext("%d heading", "%d headings", count) % count
	else:
		# Translators: Part of the page summary announced when a page has loaded.
		text = ngettext("%d link", "%d links", count) % count
	if capped:
		# Translators: Part of the page summary, used when a page has so many of
		# something that they were not all counted. {count} is e.g. "1500 links".
		text = _("over {count}").format(count=text)
	return text


def _summaryText(counts):
	phrases = {
		nodeType: _countPhrase(nodeType, *counts[nodeType])
		for nodeType in _SUMMARY_TYPES
	}
	# Translators: Announced when a web page has finished loading.
	summary = _("Page has {regions}, {headings} and {links}").format(
		regions=phrases["landmark"],
		headings=phrases["heading"],
		links=phrases["link"],
	)
	if not getAnnounceLoadingComplete():
		return summary
	# Translators: Announced when a web page has finished loading, before the summary.
	return "%s\n%s" % (_("Loading complete"), summary)


def _announceSummary(buffer):
	"""Say what has just loaded, so long as this is a browser and not Outlook."""
	global _summaryTimer
	_summaryTimer = None
	mode = getSummaryMode()
	if mode == SUMMARY_NORMAL or not getattr(buffer, _ARMED_ATTR, False):
		return
	if not _isFocusedBuffer(buffer):
		# The user left in the last fraction of a second. Leave it armed, so the
		# document load event still coming can announce it if they come back.
		return
	if not getattr(buffer, "isReady", False):
		return
	if _bufferTextLength(buffer) == 0:
		# NVDA is still waiting for content; it will report the document again from
		# event_documentLoadComplete, and this stays armed for that.
		return

	root = getattr(buffer, "rootNVDAObject", None)
	if outlookIsCurrent(root):
		log.debug("Mute Browse Mode: Outlook is the current program, no page summary")
		setattr(buffer, _ARMED_ATTR, False)
		return
	if not _isWebBrowser(root):
		log.debug("Mute Browse Mode: not a browser window, no page summary")
		setattr(buffer, _ARMED_ATTR, False)
		return
	setattr(buffer, _ARMED_ATTR, False)
	if _sayAllRunning():
		# The user asked for the whole page to be read. Do not talk over it.
		return
	if mode == SUMMARY_TONES:
		_playReadyTones()
		return
	with _ownSpeech():
		ui.message(_summaryText(_countElements(buffer)))


### Arriving in Outlook

#: How long after Outlook comes to the front we keep waiting for the focus to settle
#: on something worth describing.
_OUTLOOK_ARRIVAL_WINDOW = 3.0
#: A container with more children than this is never scanned for its selected item.
#: Outlook message lists run to thousands of rows, and NVDA must not be held up.
_MAX_CHILDREN_SCANNED = 100

#: Roles the focus passes through while an application is still coming to the front.
#: Landing on one of these means Outlook has not settled yet, so keep waiting.
_TRANSIENT_FOCUS_ROLES = _members(controlTypes.Role, ("WINDOW", "PANE", "FRAME", "APPLICATION"))

#: Roles that name the field an item sits in: the "message list" an Outlook message is
#: a row of, the folder tree a folder is in, and so on.
_CONTAINER_ROLES = _members(
	controlTypes.Role,
	(
		"LIST",
		"TREEVIEW",
		"TABLE",
		"DATAGRID",
		"DATAITEM",
		"TABCONTROL",
		"TOOLBAR",
		"GROUPING",
		"PROPERTYPAGE",
		"COMBOBOX",
	),
)


def _labelledContainerOf(obj):
	"""The named field C{obj} sits in, if it sits in one.

	Walks up only a few levels and stops at the window, so this cannot wander off into
	the top of the application and read out a window title we have just silenced.
	"""
	parent = obj
	for _step in range(4):
		try:
			parent = parent.parent
		except Exception:
			return None
		if parent is None:
			return None
		role = getattr(parent, "role", None)
		if role in _TRANSIENT_FOCUS_ROLES or role in _TITLE_ROLES:
			return None
		if role in _CONTAINER_ROLES and (getattr(parent, "name", "") or ""):
			return parent
	return None


def _selectedItemName(obj):
	"""The name of the selected child of C{obj}, when C{obj} is a small container.

	Only for the case where the focus is the container itself. Where NVDA focuses the
	item instead, which is what Outlook does with its message list and folder tree,
	the item is the focus object and this is not needed.
	"""
	try:
		if obj.role not in _CONTAINER_ROLES:
			return None
		count = obj.childCount
	except Exception:
		return None
	if not count or count > _MAX_CHILDREN_SCANNED:
		return None
	selected = getattr(controlTypes.State, "SELECTED", None)
	if selected is None:
		return None
	try:
		for child in obj.children:
			if selected in child.states and child.name:
				return child.name
	except Exception:
		return None
	return None


def _briefFocusSpeech(obj):
	"""A short description of the focus: the field, what it is, and what is in it."""
	getProperties = getattr(speech, "getObjectPropertiesSpeech", None)
	if getProperties is None:
		return []
	reason = getattr(controlTypes.OutputReason, "QUERY", None)
	sequence = []
	container = _labelledContainerOf(obj)
	if container is not None:
		sequence.extend(getProperties(container, reason=reason, name=True, role=True))
	sequence.extend(
		getProperties(
			obj,
			reason=reason,
			name=True,
			role=True,
			value=True,
			states=True,
			positionInfo_indexInGroup=True,
			positionInfo_similarItemsInGroup=True,
		),
	)
	name = _selectedItemName(obj)
	if name:
		# Translators: Part of the brief report when switching to Microsoft Outlook.
		sequence.append(_("{item} selected").format(item=name))
	return sequence


def _speakBriefFocus(sequence):
	"""Say where the user has landed, past the add-on's own gate."""
	with _ownSpeech():
		speech.speak(sequence)


### The Outlook message body

#: The window class of the Word editing surface Outlook composes a message in, and the
#: only one that is a message body.
#:
#: Deliberately not "RichEdit20W": NVDA's own Outlook support treats every window whose
#: class starts with that as a contact edit field, which is what To, Cc and Subject are,
#: so matching on it announced the message body for all of them.
#:
#: Deliberately narrow for the other direction too. Word puts an editing surface inside
#: its dialogs as well, which NVDA has a class of its own for, ``WordDocument_WwN``, and
#: in Outlook those pick up exactly the same message body markings as the real thing.
#: The box showing the misspelled word in the F7 spelling dialog is one of them, which
#: is why it announced the message body when it took the focus. Requiring the editing
#: surface itself keeps every one of them out, whatever window class it turns out to
#: have, because none of them can be the document window.
_OUTLOOK_BODY_WINDOW_CLASSES = frozenset(("_WwG",))

#: The one RichEdit control in Outlook that is a message body rather than a field:
#: the plain text message, which NVDA identifies by exactly this class and id.
_PLAIN_TEXT_BODY_CLASS = "RichEdit20W"
_PLAIN_TEXT_BODY_CONTROL_ID = 8224

#: Names the message body goes by where nothing else identifies it, used together with
#: several other tests rather than on its own.
_BODY_NAMES = frozenset(("message", "message body"))

#: Roles a message body can have. Everything else the user tabs through on the way to
#: it, To, Cc, Subject and the toolbars, has some other role.
_BODY_ROLES = _members(controlTypes.Role, ("DOCUMENT", "EDITABLETEXT"))

_STATE_READONLY = getattr(controlTypes.State, "READONLY", None)
_STATE_UNAVAILABLE = getattr(controlTypes.State, "UNAVAILABLE", None)
_STATE_MULTILINE = getattr(controlTypes.State, "MULTILINE", None)


def _hasState(states, state):
	return state is not None and state in states


def _isOutlookMessageBody(obj):
	"""Whether the focus has landed in an Outlook message body that can be typed into.

	Being editable is not the test. The address and subject fields are editable text
	too, and so is nearly everything else on a message form, which is why anything
	that loose ends up announcing the message body for the lot of them. What is needed
	is a positive identification of the body itself, and there are three:

	* NVDA's own Outlook support puts C{isReadonlyViewer} on the message body object,
	  and its value says whether this is a message being written or one being read.
	  It is not enough on its own, because Word hands the same kind of object to its
	  dialogs, so the object also has to be the editing surface itself;
	* that editing surface has a window class of its own, which no field on the form
	  and no Word dialog shares;
	* the plain text body is the one RichEdit control NVDA picks out by control id.

	A body the user cannot type into, such as the one in the reading pane, is never
	announced, because the announcement invites them to type.

	Whether this is Outlook at all is asked of the window rather than of the object: the
	web view branch below exists for an Outlook built on one, and everything inside a web
	view belongs to another process entirely. See L{_isInOutlookWindow}.
	"""
	if not _isInOutlookWindow(obj):
		return False
	if getattr(obj, "role", None) not in _BODY_ROLES:
		return False
	try:
		states = set(obj.states or ())
	except Exception:
		states = set()
	if _hasState(states, _STATE_READONLY) or _hasState(states, _STATE_UNAVAILABLE):
		return False

	windowClass = _windowClassOf(obj)
	viewer = getattr(obj, "isReadonlyViewer", None)
	if viewer is not None:
		# A Word surface Outlook has marked up as a message. Only the editing surface
		# itself is the body; the same marking is on the one inside the F7 spelling
		# dialog and every other Word dialog.
		return not viewer and windowClass in _OUTLOOK_BODY_WINDOW_CLASSES
	if windowClass in _OUTLOOK_BODY_WINDOW_CLASSES:
		return True
	if (
		windowClass == _PLAIN_TEXT_BODY_CLASS
		and getattr(obj, "windowControlID", None) == _PLAIN_TEXT_BODY_CONTROL_ID
	):
		return True
	# Outlook built on a web view has none of the above. Take a name that says it is
	# the body, but only along with taking more than one line, which no address or
	# subject field does.
	if not _hasState(states, _STATE_MULTILINE):
		return False
	return (getattr(obj, "name", "") or "").strip().lower() in _BODY_NAMES


def _announceMessageBody():
	# Translators: Announced when the focus reaches the message body in Microsoft Outlook.
	with _ownSpeech():
		ui.message(_("You are now in the message body, type a message."))


### The Outlook spelling checker
#
# The F7 window in the classic Outlook is not a modern accessibility interface at all.
# It is an ordinary Win32 dialog put on screen by the Word engine that renders the
# message, and the box in it that shows the mistake is a small Word editing surface.
# Word also highlights the mistake in the message itself as the dialog steps through
# them, so the message's own selection is a second place the word can be read from.
#
# NVDA already knows this window: the box has a window class of ``_WwN`` or ``_WwO`` and
# the control id 18, and NVDA answers those three facts with a class called
# ``SpellCheckErrorField`` whose ``errorText`` is the mistake and nothing else. The
# add-on used to look for ``_WwN`` alone and read the box's *selection*, which is why it
# could come away with nothing at all and then say nothing at all, having already
# silenced NVDA's own reading of the box. Every one of those places is now asked in turn.

#: The window classes Word gives an editing surface it puts inside one of its own
#: dialogs, which NVDA has a class of its own for, ``WordDocument_WwN``. In Outlook the
#: one the user meets is the box in the F7 spelling dialog. Both classes are here because
#: NVDA accepts either, and looking for only the first is one of the reasons the box was
#: never found.
_WORD_DIALOG_WINDOW_CLASSES = frozenset(("_WwN", "_WwO"))

#: The control id Word gives the "Not in Dictionary" box, and the one thing that tells it
#: apart from the other Word surfaces the dialog is built out of. NVDA uses exactly this
#: number to decide a window is a spelling error field rather than an ordinary document,
#: and a window with it is always the right one to ask.
_SPELL_ERROR_CONTROL_ID = 18

#: How much slower than usual the word is said and spelled out. 0.8 is a fifth slower.
_SPELL_RATE_MULTIPLIER = 0.8

#: Longer than this and whatever we found is not one misspelled word.
_MAX_WORD_LENGTH = 60

#: How long to wait before looking again for the word the dialog is asking about, in
#: milliseconds, and how many times. Word puts the dialog on screen before it has
#: selected the error in the message, so the first look — which happens as the dialog
#: takes the focus — can come up empty, and the first word of a check was announced as
#: nothing at all. They stop as soon as a word is found or the focus leaves the dialog.
_SPELL_RETRY_DELAYS = (120, 300, 600)

#: Punctuation taken off either end of what the checker has selected. The selection
#: routinely takes in the full stop or the comma the error is sitting against, and the
#: apostrophes and hyphens inside a word are left alone because they are part of it.
_WORD_EDGE_PUNCTUATION = " \t\r\n.,;:!?\"'`()[]{}<>…«»„“”‘’-–—/\\|*_+=@#$%^&~"


def _looksLikeWord(text):
	"""Whether C{text} is one word, rather than a sentence or nothing at all.

	Punctuation on its own is not a word. Word keeps a selection in the message as the
	dialog steps along, and at the end of a check that selection can be sitting on
	nothing but the full stop the last sentence finished with — which is how the add-on
	came to announce "Not in Dictionary: period". A misspelling has letters in it, so
	that is what is asked for.
	"""
	word = (text or "").strip()
	if not word or len(word) > _MAX_WORD_LENGTH:
		return False
	if any(character.isspace() for character in word):
		return False
	return any(character.isalpha() for character in word)


try:
	# NVDA declares both of these with the right argument types, so this is its own
	# binding rather than a second one built beside it.
	from winBindings.user32 import WNDENUMPROC as _ENUM_WINDOWS_PROC
	from winBindings.user32 import EnumChildWindows as _enumChildWindows
except Exception:
	try:
		_ENUM_WINDOWS_PROC = ctypes.WINFUNCTYPE(
			ctypes.wintypes.BOOL,
			ctypes.wintypes.HWND,
			ctypes.wintypes.LPARAM,
		)
		_enumChildWindows = ctypes.windll.user32.EnumChildWindows
		_enumChildWindows.restype = ctypes.wintypes.BOOL
		_enumChildWindows.argtypes = (
			ctypes.wintypes.HWND,
			_ENUM_WINDOWS_PROC,
			ctypes.wintypes.LPARAM,
		)
	except Exception:
		# Without it the spelling box is looked for the way NVDA looks for one, which
		# finds it in every ordinary case; only a dialog holding more than one Word
		# surface needs the walk below to tell them apart.
		_ENUM_WINDOWS_PROC = None
		_enumChildWindows = None
		log.debugWarning("Mute Browse Mode: cannot walk child windows", exc_info=True)

#: Nothing sane has this many windows inside it. A ceiling so a walk of the window tree
#: can never become the slow thing in a focus change.
_MAX_WINDOWS_WALKED = 400


def _childWindows(window):
	"""Every window inside C{window}, however deeply nested, or an empty list."""
	found = []
	if not window or _enumChildWindows is None:
		return found

	def visit(child, _lParam):
		found.append(child)
		return len(found) < _MAX_WINDOWS_WALKED

	try:
		_enumChildWindows(window, _ENUM_WINDOWS_PROC(visit), 0)
	except Exception:
		log.debugWarning("Mute Browse Mode: could not walk the dialog", exc_info=True)
	return found


def _anyWordDialogSurface(obj):
	"""Any Word dialog surface in the same thread as C{obj}, or 0 for none.

	NVDA's own search, in one pass of native code, and the whole question for every focus
	change where the spelling window is not open — which is nearly all of them. Only when
	this says there is a Word dialog somewhere does it become worth looking properly.
	"""
	try:
		import NVDAHelper

		for windowClass in sorted(_WORD_DIALOG_WINDOW_CLASSES):
			window = NVDAHelper.localLib.findWindowWithClassInThread(
				obj.windowThreadID,
				windowClass,
				True,
			)
			if window:
				return window
	except Exception:
		log.debugWarning("Mute Browse Mode: could not look for a Word dialog", exc_info=True)
	return 0


def _spellBoxInside(window):
	"""The box showing the mistake, somewhere inside C{window}, or 0.

	A Word dialog is built out of several Word surfaces and only one of them shows the
	mistake, so they are told apart by control id rather than by which turns up first.
	Any of them is better than none, so an unrecognised one is kept as a spare.
	"""
	spare = 0
	for child in _childWindows(window):
		try:
			if winUser.getClassName(child) not in _WORD_DIALOG_WINDOW_CLASSES:
				continue
			if winUser.getControlID(child) == _SPELL_ERROR_CONTROL_ID:
				return child
		except Exception:
			continue
		spare = spare or child
	return spare


def _sameWindowFamily(window, obj):
	"""Whether C{window} and C{obj} belong to the same family of windows.

	"The same window" is the wrong test for a dialog, because a dialog is a window in its
	own right rather than a part of the window it belongs to. What the two share is the
	window that owns them both, which for the spelling dialog is the message being
	checked.
	"""
	try:
		owner = getattr(winUser, "GA_ROOTOWNER", 3)
		return winUser.getAncestor(window, owner) == winUser.getAncestor(obj.windowHandle, owner)
	except Exception:
		return False


def _spellErrorFieldWindow(obj):
	"""The handle of the box showing the mistake, in the spelling dialog C{obj} is in.

	The dialog is looked for in the window the focus is in, in the window in front, and
	in the windows either of those owns, and nowhere else: a Word dialog open somewhere
	else in Outlook is not the one being looked at.
	"""
	if _windowClassOf(obj) in _WORD_DIALOG_WINDOW_CLASSES:
		try:
			if winUser.getControlID(obj.windowHandle) == _SPELL_ERROR_CONTROL_ID:
				return obj.windowHandle
		except Exception:
			pass
	surface = _anyWordDialogSurface(obj)
	if not surface:
		return 0
	try:
		containers = []
		for container in (_rootWindowOf(obj), winUser.getForegroundWindow()):
			if container and container not in containers:
				containers.append(container)
		for container in containers:
			window = _spellBoxInside(container)
			if window:
				return window
		# The dialog is not part of either of those, so it is a window of its own. Take
		# the one NVDA's own search turned up, so long as it belongs to the same family
		# of windows, and look inside it for the box proper.
		found = winUser.getAncestor(surface, winUser.GA_ROOT)
		if found in containers or _sameWindowFamily(surface, obj):
			return _spellBoxInside(found) or surface
	except Exception:
		log.debugWarning("Mute Browse Mode: could not look for the spelling dialog", exc_info=True)
	return 0


def _objectFromWindow(window):
	"""A fresh NVDA object for C{window}, or C{None}.

	Fresh matters: NVDA caches an object's properties, and the whole point of asking
	again a moment later is to get an answer that has changed since.
	"""
	try:
		from NVDAObjects.IAccessible import getNVDAObjectFromEvent

		return getNVDAObjectFromEvent(window, winUser.OBJID_CLIENT, 0)
	except Exception:
		log.debugWarning("Mute Browse Mode: could not reach the spelling box", exc_info=True)
		return None


def _outlookSpellCheckField(obj):
	"""The box showing the mistake, in the spelling dialog the focus is in, if it is.

	The focus lands on whichever control the dialog starts on, and moves about as the
	user works, so the box is looked for rather than assumed. See
	L{_spellErrorFieldWindow} for how it is picked out.
	"""
	if not _isInOutlookWindow(obj):
		return None
	window = _spellErrorFieldWindow(obj)
	if not window:
		if _tracing and _windowClassOf(obj) in _WORD_DIALOG_WINDOW_CLASSES:
			# The focus is sitting on a Word dialog surface and the box was still not
			# found, which is the one thing worth knowing if this ever goes quiet again.
			_traceWrite("SPELL no box found from %s" % _describe(obj))
		return None
	if _tracing:
		try:
			_traceWrite(
				"SPELL box=%s class=%s id=%s"
				% (window, winUser.getClassName(window), winUser.getControlID(window)),
			)
		except Exception:
			_traceWrite("SPELL box=%s" % window)
	if getattr(obj, "windowHandle", None) == window:
		return obj
	return _objectFromWindow(window)


def _selectionText(obj):
	"""What is selected in C{obj}, or what the cursor is on if nothing is."""
	for position, unit in (
		(textInfos.POSITION_SELECTION, None),
		(textInfos.POSITION_CARET, textInfos.UNIT_WORD),
	):
		try:
			info = obj.makeTextInfo(position)
			if unit is not None:
				info.expand(unit)
			text = info.text or ""
		except Exception:
			continue
		if text.strip():
			return text
	return None


def _boldRun(field):
	"""The bold part of the box, which is how the dialog shows which word it means.

	The box holds the whole sentence the mistake sits in and shows the mistake itself in
	bold. This is how NVDA reads the word out of older versions of Word, and it works
	whether or not Word will answer questions about itself.
	"""
	from displayModel import EditableTextDisplayModelTextInfo

	info = EditableTextDisplayModelTextInfo(field, textInfos.POSITION_ALL)
	inBold = False
	bold = []
	for item in info.getTextWithFields():
		if isinstance(item, str):
			if inBold:
				bold.append(item)
		elif getattr(item, "field", None):
			inBold = item.field.get("bold", False)
		if not inBold and bold:
			break
	return "".join(bold)


def _messageSelection(field):
	"""What Word has selected in the message itself.

	The dialog highlights each mistake in the email as it steps through them, so the
	message's own selection is the word being asked about. Asked of the message rather
	than of the dialog, so it is an answer even when the box will not give one.
	"""
	import NVDAHelper

	for windowClass in sorted(_OUTLOOK_BODY_WINDOW_CLASSES):
		window = NVDAHelper.localLib.findWindowWithClassInThread(
			field.windowThreadID,
			windowClass,
			True,
		)
		if not window:
			continue
		body = _objectFromWindow(window)
		if body is None:
			continue
		text = _selectionText(body)
		if text:
			return text
	return None


def _wordSources(field):
	"""Every way of asking what word the checker has stopped on, best answer first.

	Each is a name for the trace and something to call. They are in this order because
	the first two are what NVDA itself reads, are the word on its own rather than the
	sentence around it, and cost nothing; the rest are for a dialog that does not answer
	the way NVDA expects, and end with asking the email instead of the dialog.
	"""
	return (
		("errorText", lambda: getattr(field, "errorText", None)),
		("value", lambda: getattr(field, "value", None)),
		("bold", lambda: _boldRun(field)),
		("selection", lambda: _selectionText(field)),
		("message", lambda: _messageSelection(field)),
	)


def _misspelledWord(field):
	"""The word the spelling checker is asking about, or C{None}.

	Whatever comes back has to look like a single word, or it is not used at all: several
	of the places asked hold the whole sentence the mistake is in, and reading that out
	would be worse than saying nothing.
	"""
	try:
		field.invalidateCache()
	except Exception:
		pass
	for name, source in _wordSources(field):
		try:
			word = (source() or "").strip().strip(_WORD_EDGE_PUNCTUATION)
		except Exception:
			if _tracing:
				_traceWrite("SPELL %s raised" % name)
			continue
		if _looksLikeWord(word):
			if _tracing:
				_traceWrite("SPELL %s gave %r" % (name, word))
			return word
		if _tracing and word:
			_traceWrite("SPELL %s gave %r, not one word" % (name, word[:80]))
	return None


def _spellingSpeech(word):
	""""Misspelled", then the word said and spelled out a fifth more slowly than usual.

	The shape JAWS uses: you are told that this is a word the checker does not know, and
	then you are given it, twice — as a word and letter by letter — slowly enough to take
	in. Only the word slows down; the label in front of it is said at the usual speed,
	because it is the same label every time and is not the part to listen to.

	The slower rate belongs to this one announcement rather than to the synthesiser, so
	the speed goes back to normal straight afterwards and cannot be left behind anywhere
	else, not even if the spelling is interrupted half way through.
	"""
	# Translators: Said in the Outlook spelling window, in front of the word the checker
	# has stopped on.
	label = _("misspelled")
	sequence = [label, word]
	spoken = [label]
	try:
		if _EndUtteranceCommand is not None:
			spoken.append(_EndUtteranceCommand())
		if _RateCommand is not None:
			spoken.append(_RateCommand(multiplier=_SPELL_RATE_MULTIPLIER))
		spoken.append(word)
		getSpelling = getattr(speech, "getSpellingSpeech", None)
		if getSpelling is not None:
			if _EndUtteranceCommand is not None:
				spoken.append(_EndUtteranceCommand())
			spoken.extend(getSpelling(word))
		if _RateCommand is not None:
			# Back to the rate the user chose. The synthesiser is never reconfigured, so
			# this cannot escape the one announcement even if it is interrupted.
			spoken.append(_RateCommand())
	except Exception:
		# Say the word even if it cannot be spelled out or slowed down. Half an
		# announcement is worth having; the dialog arriving in silence is not.
		log.debugWarning("Mute Browse Mode: could not spell the word out", exc_info=True)
		return sequence
	return spoken


def _announceMisspelledWord(word):
	try:
		sequence = _spellingSpeech(word)
	except Exception:
		log.error("Mute Browse Mode: could not build the spelling announcement", exc_info=True)
		sequence = [word]
	with _ownSpeech():
		speech.speak(sequence)


### The dialog that says the check has finished
#
# A check ends with a small dialog of its own holding a single OK button, and it comes up
# whether or not anything was found. A check that found nothing at all puts it straight on
# screen without a word ever having been said, so there is nothing behind it to match it
# against, and matching it against the dialog the words were asked in does not work either
# — it is a window in its own right, so it does not share one with anything. What is left
# is what it says, and it says so plainly: "The spelling and grammar check is complete."
#
# The dialog's own words are read rather than assumed, and two things have to be in them:
# something about spelling, and something about being finished. Either on its own is not
# enough. "Text marked with do not check spelling or grammar was skipped" is about
# spelling and is not this dialog; plenty of Outlook dialogs say something is complete.

#: The names the OK button goes by. The ampersand is the underline under the letter you
#: can press with alt, which some versions leave in the name.
_OK_BUTTON_NAMES = frozenset(("ok", "&ok"))

#: Roles a dialog has. The walk up from the button stops at the first of these, and gives
#: up at anything that is a whole window, so that failing to find the dialog can never
#: turn into reading the whole of Outlook.
_DIALOG_ROLES = _members(controlTypes.Role, ("DIALOG", "ALERT", "PROPERTYPAGE", "OPTIONPANE"))
_NOT_A_DIALOG_ROLES = _members(controlTypes.Role, ("APPLICATION", "FRAME", "DESKTOP"))

#: Window classes that are a dialog whatever role NVDA gives them: the standard Windows
#: one, and the one Office uses for its own.
_DIALOG_WINDOW_CLASS = "#32770"
_OFFICE_DIALOG_WINDOW_PREFIX = "bosa_sdm"

#: How far up from the button to look for the dialog, and how much of the dialog to read
#: once it is found. A message box is a title, a line of text and a button or two, so
#: these are generous; they are here so that an unexpected shape cannot turn a focus
#: change into a walk of something enormous.
_DIALOG_WALK_LIMIT = 6
_DIALOG_SCAN_DEPTH = 4
_DIALOG_SCAN_LIMIT = 60


def _wordList(english, translated):
	"""The English words and the translated ones, since the dialog is not NVDA's.

	NVDA may be running in one language and Outlook in another, so the English words are
	always looked for as well as the ones for the user's language.
	"""
	words = set()
	for source in (english, translated):
		for word in (source or "").split(","):
			word = word.strip().lower()
			if word:
				words.add(word)
	return frozenset(words)


#: Translators: Parts of words, separated by commas and in lower case, that say a message
#: is about spelling. They are matched against what the Outlook dialog announcing the end
#: of a spelling check says, so they should be the shortest stem that cannot be mistaken
#: for something else: "spell" rather than "spelling", so that "spellcheck" matches too.
_SPELLING_WORDS = _wordList("spell,grammar,proofing", _("spell,grammar,proofing"))

#: Translators: Parts of words, separated by commas and in lower case, that say something
#: has finished. Matched against the same dialog as the list above, and both lists have to
#: match before the add-on will say a spelling check is complete.
_COMPLETION_WORDS = _wordList("complete,finish,done", _("complete,finish,done"))


def _isOkButton(obj):
	"""Whether C{obj} is an OK button in Outlook."""
	if not _isInOutlookWindow(obj):
		return False
	try:
		if getattr(obj, "role", None) != controlTypes.Role.BUTTON:
			return False
		return (getattr(obj, "name", "") or "").strip().lower() in _OK_BUTTON_NAMES
	except Exception:
		return False


def _looksLikeADialog(obj):
	"""Whether C{obj} is the dialog itself rather than something on the way to it."""
	if getattr(obj, "role", None) in _DIALOG_ROLES:
		return True
	windowClass = _windowClassOf(obj)
	return windowClass == _DIALOG_WINDOW_CLASS or windowClass.lower().startswith(
		_OFFICE_DIALOG_WINDOW_PREFIX,
	)


def _dialogAbove(obj):
	"""The dialog C{obj} sits in, or C{None} if it is not in one.

	C{None} matters as much as an answer here. The whole point of reading a dialog is to
	find out what it is, so reading the wrong thing is worse than reading nothing: the
	walk gives up rather than settling for whatever it has reached.
	"""
	current = obj
	for _step in range(_DIALOG_WALK_LIMIT):
		try:
			current = current.parent
		except Exception:
			return None
		if current is None:
			return None
		if _looksLikeADialog(current):
			return current
		if getattr(current, "role", None) in _NOT_A_DIALOG_ROLES:
			# Out of the dialog and into the application without having found one.
			return None
	return None


def _dialogText(dialog):
	"""Everything C{dialog} says, in one lower case string, or C{None}.

	A message box keeps its message in a piece of static text inside it rather than in
	its own name, so the children are read as well as the dialog. Bounded twice over, by
	how deep it goes and by how many objects it holds, because this runs on a focus
	change.

	Running past that bound is an answer in itself, and the answer is C{None}. The dialog
	being looked for is a message box: a title, a line of text and a button. Anything with
	more in it than that is something else — an Outlook message is itself a dialog — and
	reading a whole message and finding the words "spelling" and "complete" somewhere in
	the middle of it is exactly the mistake this is here to avoid.
	"""
	texts = []
	read = 0
	level = [dialog]
	for _depth in range(_DIALOG_SCAN_DEPTH):
		below = []
		for item in level:
			read += 1
			if read > _DIALOG_SCAN_LIMIT:
				return None
			for name in ("name", "value", "description"):
				try:
					text = getattr(item, name, None)
				except Exception:
					continue
				if isinstance(text, str) and text.strip():
					texts.append(text)
			try:
				below.extend(item.children or ())
			except Exception:
				pass
			if len(below) > _DIALOG_SCAN_LIMIT:
				return None
		if not below:
			break
		level = below
	return " ".join(texts).lower()


def _saysSpellCheckIsComplete(obj):
	"""Whether the dialog C{obj} is in says a spelling check has finished."""
	dialog = _dialogAbove(obj)
	if dialog is None:
		return False
	try:
		text = _dialogText(dialog)
	except Exception:
		log.debugWarning("Mute Browse Mode: could not read the dialog", exc_info=True)
		return False
	if not text:
		return False
	aboutSpelling = any(word in text for word in _SPELLING_WORDS)
	hasFinished = any(word in text for word in _COMPLETION_WORDS)
	if _tracing:
		_traceWrite(
			"SPELLDONE spelling=%s finished=%s text=%r"
			% (aboutSpelling, hasFinished, text[:160]),
		)
	return aboutSpelling and hasFinished


def _spellCheckCompleteSpeech(onTheOkButton):
	"""What to say when a check ends.

	The button is only named where the focus is actually on it. A check can also end with
	Outlook simply handing the message back, and naming a button that is not there would
	send the user looking for one.
	"""
	sequence = [
		# Translators: Said in place of NVDA's own report when the Outlook spelling
		# checker has finished and its last dialog, holding only an OK button, comes up.
		_("Spell check is complete."),
	]
	if onTheOkButton:
		# Translators: The button on that dialog, named after the sentence above it.
		sequence.append(_("OK button"))
	return sequence


#: What NVDA calls the state a disabled control is in, lower case, or C{None} where this
#: NVDA does not name its states this way.
try:
	_UNAVAILABLE_TEXT = (_STATE_UNAVAILABLE.displayString or "").strip().lower() or None
except Exception:
	_UNAVAILABLE_TEXT = None
	log.debugWarning("Mute Browse Mode: no name for the unavailable state", exc_info=True)


def _isBareUnavailable(args, kwargs):
	"""Whether a whole utterance is nothing but the word "unavailable".

	Word disables the spelling dialog the moment the check is complete, and NVDA answers
	that state change on whatever had the focus with that one word — immediately before
	the add-on says the check has finished, which is why the user hears "unavailable,
	spell check is complete". It says nothing the user can act on, and it is only ever
	dropped while a check this add-on is following is actually in progress.
	"""
	if _UNAVAILABLE_TEXT is None:
		return False
	if args:
		sequence = args[0]
	elif "speechSequence" in kwargs:
		sequence = kwargs["speechSequence"]
	else:
		return False
	if not isinstance(sequence, list) or not sequence:
		return False
	text = " ".join(item for item in sequence if isinstance(item, str))
	return text.strip().strip(_WORD_EDGE_PUNCTUATION).lower() == _UNAVAILABLE_TEXT


### The progress bar
#
# Outlook shows one while it fetches new mail, and NVDA reads it out a percentage at a
# time: "ten percent", "twenty percent", on to a hundred. Those announcements were being
# swallowed part of the way through. Each message that lands opens a document, every
# document announcement holds the gate shut behind it, and the download carried on
# silently underneath — so it would start being read out and then stop, with nothing to
# say whether it had finished or given up.
#
# A progress bar is never the noise this add-on was written to remove. It is the opposite:
# something is happening that the user cannot see and is waiting on. So it goes past the
# gate, always.
#
# Getting it past takes one step more than usual, because NVDA does not speak it where it
# decides to. ``behaviors.ProgressBar.event_valueChange`` *queues* ``speech.speakMessage``
# on the event queue and returns, so wrapping that call in ``_ownSpeech`` would be over
# and done with an event-queue turn before anything was said. What is wrapped instead is
# the function being queued, so the bypass travels on the queue along with it.

#: The stand-in for ``speech.speakMessage`` and the function it stands in for, kept so it
#: is built once rather than on every tick of every progress bar.
_progressSpeakMessage = (None, None)


def _bypassingSpeakMessage(original):
	"""``speech.speakMessage``, but never gated."""
	global _progressSpeakMessage
	if _progressSpeakMessage[0] is original:
		return _progressSpeakMessage[1]

	def speakMessage(*args, **kwargs):
		with _ownSpeech():
			return original(*args, **kwargs)

	speakMessage.__name__ = "speakMessage"
	speakMessage.__doc__ = getattr(original, "__doc__", None)
	_progressSpeakMessage = (original, speakMessage)
	return speakMessage


def _describeProgress(obj):
	"""A progress bar's value and whether NVDA counts it as being in front, for the trace.

	Those two are the whole of why NVDA might say nothing about it of its own accord: a
	bar that is not in the foreground is only reported when "report background progress
	bars" is turned on in NVDA's own settings, which it is not by default.
	"""
	parts = []
	for label, getter in (
		("value", lambda: repr(obj.value)),
		("inForeground", lambda: str(obj.isInForeground)),
		("app", lambda: _appNameOf(obj)),
	):
		try:
			parts.append("%s=%s" % (label, getter()))
		except Exception:
			parts.append("%s=?" % label)
	return " ".join(parts)


def _makeProgressBarWrapper(original):
	"""``ProgressBar.event_valueChange``, with what it queues let past the gate.

	The swap lasts for the length of one call on NVDA's main thread, which is where every
	value change is handled, and it is undone whether or not that call succeeds. Nothing
	else in NVDA can see it: the only thing that reads ``speech.speakMessage`` in between
	is the call this is wrapping.
	"""

	def event_valueChange(self, *args, **kwargs):
		if _tracing:
			# Worth a line of its own: it is the one thing that tells "NVDA never
			# offered a percentage" apart from "NVDA offered one and it was silenced".
			_traceWrite("PROGRESS %s" % _describeProgress(self))
		current = getattr(speech, "speakMessage", None)
		if current is None or getMode() == MODE_NORMAL:
			# Nothing is being silenced, so there is nothing to let past.
			return original(self, *args, **kwargs)
		speech.speakMessage = _bypassingSpeakMessage(current)
		try:
			return original(self, *args, **kwargs)
		finally:
			speech.speakMessage = current

	event_valueChange.__name__ = "event_valueChange"
	event_valueChange.__doc__ = getattr(original, "__doc__", None)
	return event_valueChange


### Putting links on their own line in Outlook

# NVDA already does this in a web browser: Browse Mode settings, "Use screen layout
# (when supported)". With it off, a link or a button that shares a line with other
# things on screen gets a line of its own in browse mode, so down arrow reaches each of
# them in turn. The setting is read in exactly one place, VirtualBufferTextInfo._
# getLineOffsets, and passed straight to the code that works out where a line starts
# and ends, so an Outlook message that is rendered as a web document can be given the
# same treatment on its own without touching the setting the browsers use.
#
# An Outlook message that Word renders instead is not a virtual buffer at all, and NVDA
# has no equivalent for it: the base implementation of the command that toggles this
# answers "not supported in this document". Nothing can be done there from here, so it
# says so in the log rather than quietly doing nothing.

#: The controls that earn a line of their own. Text either side of one gets a line of
#: its own too, which is what makes the link the only thing on its line.
_SPLIT_CONTROL_ROLES = _members(
	controlTypes.Role,
	(
		"LINK",
		"BUTTON",
		"TOGGLEBUTTON",
		"MENUBUTTON",
		"DROPDOWNBUTTON",
		"SPLITBUTTON",
		"CHECKBOX",
		"RADIOBUTTON",
		"COMBOBOX",
		"EDITABLETEXT",
		"SLIDER",
		"SPINBUTTON",
	),
)


def _isVirtualBufferOutlookMessage(textInfo):
	"""Whether this text is an Outlook message rendered as a web document.

	Which is Outlook's own web view, so the document belongs to msedgewebview2.exe and
	only the window it is in says Outlook. See L{_isInOutlookWindow}.
	"""
	buffer = getattr(textInfo, "obj", None)
	if buffer is None or not getattr(buffer, "VBufHandle", None):
		return False
	return _isInOutlookWindow(getattr(buffer, "rootNVDAObject", None))


def _shouldSplitLines(textInfo):
	"""Whether this line should be broken up so each control gets one of its own."""
	if not getLinksOnOwnLine():
		return False
	try:
		if not config.conf["virtualBuffers"]["useScreenLayout"]:
			# NVDA is already doing it everywhere, so there is nothing to add.
			return False
	except Exception:
		pass
	return _isVirtualBufferOutlookMessage(textInfo)


def _makeLineOffsetsWrapper(original):
	"""Work out a line as if screen layout were off, for Outlook messages only."""

	def _getLineOffsets(self, offset):
		try:
			split = _shouldSplitLines(self)
		except Exception:
			split = False
		if not split:
			return original(self, offset)
		try:
			import NVDAHelper

			lineStart = ctypes.c_int()
			lineEnd = ctypes.c_int()
			NVDAHelper.localLib.VBuf_getLineOffsets(
				self.obj.VBufHandle,
				offset,
				config.conf["virtualBuffers"]["maxLineLength"],
				False,
				ctypes.byref(lineStart),
				ctypes.byref(lineEnd),
			)
			return lineStart.value, lineEnd.value
		except Exception:
			log.debugWarning("Mute Browse Mode: could not split the line", exc_info=True)
			return original(self, offset)

	return _getLineOffsets


### ...and the same thing where there is no buffer to do it for us
#
# An Outlook message that Word renders is not a virtual buffer, and NVDA has no screen
# layout for it: the base implementation of the command that toggles that answers "not
# supported in this document". Its lines are Word's own, so a link sitting in the middle
# of a sentence is read out as part of that sentence and down arrow steps straight over
# it.
#
# Redefining what a line is would reach much too far: NVDA uses the line unit for
# braille, for reporting the line the focus lands on, and for say all. So instead only
# the down and up arrow scripts are wrapped, and only for this one kind of document.
# They walk the line in segments, split where a control starts and ends, so the link is
# on its own and the words either side of it are too. Everything else about the document
# is left exactly as NVDA has it, and any difficulty falls straight back to NVDA's own
# line movement.


def _shouldWalkSegments(treeInterceptor):
	"""Whether down and up arrow should step through this document in segments."""
	if not getLinksOnOwnLine():
		return False
	if not isinstance(treeInterceptor, browseMode.BrowseModeDocumentTreeInterceptor):
		return False
	if getattr(treeInterceptor, "passThrough", False):
		# Focus mode: the arrows belong to the application.
		return False
	if getattr(treeInterceptor, "VBufHandle", None):
		# A web document, where the line offsets above have already done the job.
		return False
	return _isInOutlookWindow(getattr(treeInterceptor, "rootNVDAObject", None))


def _isSplittableControl(field):
	try:
		return field.get("role") in _SPLIT_CONTROL_ROLES
	except Exception:
		return False


def _lineSegments(lineInfo):
	"""Where C{lineInfo} should be broken up, as a list of (start, end) character pairs.

	The boundaries are the places a control starts and finishes, so a line reading
	"comment this is a link more words" comes back as three segments: the words before
	the link, the link, and the words after it.
	"""
	try:
		fields = lineInfo.getTextWithFields()
	except Exception:
		log.debugWarning("Mute Browse Mode: could not read the line's controls", exc_info=True)
		return None
	offset = 0
	stack = []
	edges = set()
	for field in fields:
		if isinstance(field, str):
			offset += len(field)
			continue
		command = getattr(field, "command", None)
		if command == "controlStart":
			stack.append((offset, _isSplittableControl(getattr(field, "field", None))))
		elif command == "controlEnd" and stack:
			start, splittable = stack.pop()
			if splittable and offset > start:
				edges.add(start)
				edges.add(offset)
	total = offset
	if total <= 0:
		# A blank line has nothing to split. NVDA's own movement handles those.
		return None
	whole = [(0, total)]
	if not edges:
		return whole
	edges.update((0, total))
	bounds = sorted(edge for edge in edges if 0 <= edge <= total)
	text = lineInfo.text or ""
	segments = [
		(start, end)
		for start, end in zip(bounds, bounds[1:])
		if end > start and text[start:end].strip()
	]
	return segments or whole


def _segmentInfo(lineInfo, start, end):
	"""A copy of C{lineInfo} narrowed to the characters between C{start} and C{end}."""
	segment = lineInfo.copy()
	segment.collapse()
	if start and segment.move(textInfos.UNIT_CHARACTER, start) != start:
		return None
	finish = lineInfo.copy()
	finish.collapse()
	if end and finish.move(textInfos.UNIT_CHARACTER, end) != end:
		return None
	segment.setEndPoint(finish, "endToStart")
	return segment


def _caretOffsetInLine(lineInfo, caretInfo):
	prefix = lineInfo.copy()
	prefix.setEndPoint(caretInfo, "endToStart")
	return len(prefix.text or "")


def _walkSegment(treeInterceptor, gesture, direction):
	"""Move to the next or previous segment. False means "let NVDA do it instead"."""
	caret = treeInterceptor.makeTextInfo(textInfos.POSITION_CARET)
	line = caret.copy()
	line.expand(textInfos.UNIT_LINE)
	segments = _lineSegments(line)
	target = None
	if segments:
		here = _caretOffsetInLine(line, caret)
		index = 0
		for position, (start, _end) in enumerate(segments):
			if start <= here:
				index = position
		wanted = index + direction
		if 0 <= wanted < len(segments):
			target = _segmentInfo(line, *segments[wanted])
	if target is None:
		# Nothing left on this line, so take the next one and start at its near end.
		line = caret.copy()
		line.expand(textInfos.UNIT_LINE)
		line.collapse()
		if line.move(textInfos.UNIT_LINE, direction) == 0:
			return False
		line.expand(textInfos.UNIT_LINE)
		segments = _lineSegments(line)
		if not segments:
			return False
		target = _segmentInfo(line, *(segments[0] if direction > 0 else segments[-1]))
	if target is None:
		return False
	selection = target.copy()
	selection.collapse()
	# Spoken before the selection moves, the way NVDA does it: moving the selection can
	# move the focus, and that can leave the text we are about to speak behind.
	willResume = False
	try:
		willResume = scriptHandler.willSayAllResume(gesture)
	except Exception:
		pass
	if not willResume:
		speech.speakTextInfo(
			target,
			unit=textInfos.UNIT_LINE,
			reason=controlTypes.OutputReason.CARET,
		)
	try:
		treeInterceptor.selection = selection
	except Exception:
		# Spoken already, so the move must still count as done: letting NVDA's own
		# movement run now would say the next thing twice.
		log.error("Mute Browse Mode: could not move to the segment", exc_info=True)
	return True


def _makeMoveByLineWrapper(original, direction):
	def script_moveByLine(self, gesture):
		try:
			walk = _shouldWalkSegments(self)
		except Exception:
			walk = False
		if not walk:
			return original(self, gesture)
		try:
			if scriptHandler.isScriptWaiting():
				# Keys are backed up. NVDA drops the move rather than falling behind,
				# and so must we, or holding the arrow down would speak twice per press.
				return
			if _walkSegment(self, gesture, direction):
				return
		except Exception:
			log.error("Mute Browse Mode: could not walk the line in segments", exc_info=True)
		return original(self, gesture)

	script_moveByLine.__name__ = getattr(original, "__name__", "script_moveByLine")
	script_moveByLine.__doc__ = getattr(original, "__doc__", None)
	# resumeSayAllMode and anything else scriptHandler reads off the script.
	script_moveByLine.__dict__.update(getattr(original, "__dict__", {}))
	return script_moveByLine


### Monkey patching

#: (owner, name, original, wasOwnAttribute, replacement) for everything we patched.
_patches = []

#: Flags eventHandler and scriptHandler read off the handler they are about to call,
#: which have to survive being wrapped.
_CARRIED_FLAGS = ("ignoreIsReady",)


def _patch(owner, name, replacement):
	original = getattr(owner, name)
	wasOwn = name in getattr(owner, "__dict__", {})
	setattr(owner, name, replacement)
	_patches.append((owner, name, original, wasOwn, replacement))
	return original


def _unpatchAll():
	while _patches:
		owner, name, original, wasOwn, replacement = _patches.pop()
		try:
			if getattr(owner, name, None) is not replacement:
				# Something else patched on top of us; leave it alone rather than
				# clobbering another add-on.
				continue
			if wasOwn:
				setattr(owner, name, original)
			else:
				delattr(owner, name)
		except Exception:
			log.error("Mute Browse Mode: could not restore %s.%s" % (owner, name), exc_info=True)


def _adoptOriginal(wrapper, original, name):
	wrapper.__name__ = name
	wrapper.__doc__ = getattr(original, "__doc__", None)
	for flag in _CARRIED_FLAGS:
		if hasattr(original, flag):
			setattr(wrapper, flag, getattr(original, flag))


def _hookGate(owner, name, after, chime=False, onlyIf=None, chimeCheck=None, post=None):
	"""Silence C{owner.name} while it runs, and for a moment afterwards.

	Speech is dropped by two separate mechanisms. While the call is on the stack a
	depth counter silences everything, and that is the one that actually swallows the
	document announcement, all of which NVDA speaks inline. Once the call returns only
	the deadline gate is left, and that one stands down while a new window is being
	announced, so it can never eat the announcement of the control the user just
	switched to.

	@param after: seconds the deadline gate is held once the call returns.
	@param chime: play the ready tones on the way out, in "play tones" mode.
	@param onlyIf: optional callable(self, args, kwargs) deciding whether this
		particular call is one of the ones we silence. Checked before the original
		runs, because some of what it looks at is cleared by the original itself.
	@param chimeCheck: optional callable(self, args, kwargs) vetoing only the chime.
	@param post: optional callable(self, args, kwargs) run once the call returns,
		whatever the mute mode is. The page summary has its own setting, so it must
		not be switched off by this one.
	"""
	original = getattr(owner, name)

	def wrapper(self, *args, **kwargs):
		mode = getMode()
		silence = mode != MODE_NORMAL and (onlyIf is None or onlyIf(self, args, kwargs))
		if silence:
			_enterDocumentCall()
		try:
			return original(self, *args, **kwargs)
		finally:
			if silence:
				_exitDocumentCall()
				_openGate(after)
				if chime and mode == MODE_TONES and (chimeCheck is None or chimeCheck(self, args, kwargs)):
					_playReadyTones()
			if post is not None:
				try:
					post(self, args, kwargs)
				except Exception:
					log.error("Mute Browse Mode: %s post hook failed" % name, exc_info=True)

	_adoptOriginal(wrapper, original, name)
	_patch(owner, name, wrapper)


def _hookAfter(owner, name, post):
	"""Run C{post(self, args, kwargs)} once C{owner.name} returns, and nothing else."""
	original = getattr(owner, name)

	def wrapper(self, *args, **kwargs):
		try:
			return original(self, *args, **kwargs)
		finally:
			try:
				post(self, args, kwargs)
			except Exception:
				log.error("Mute Browse Mode: %s post hook failed" % name, exc_info=True)

	_adoptOriginal(wrapper, original, name)
	_patch(owner, name, wrapper)


def _hookSilent(owner, name):
	"""Silence C{owner.name} whenever it runs for Outlook or a Chromium window.

	The call still happens, so braille and NVDA's caches are unaffected; only what it
	would have spoken is dropped.
	"""
	original = getattr(owner, name)

	def wrapper(self, *args, **kwargs):
		if getMode() == MODE_NORMAL or _targetOf(self) is None:
			return original(self, *args, **kwargs)
		with _hardMute():
			return original(self, *args, **kwargs)

	_adoptOriginal(wrapper, original, name)
	_patch(owner, name, wrapper)


def _loadSucceeded(self, args, kwargs):
	"""_loadBufferDone(self, success=True): don't chime for a failed load."""
	if "success" in kwargs:
		return bool(kwargs["success"])
	if args:
		return bool(args[0])
	return True


def _summaryAfterLoad(buffer, args, kwargs):
	"""_loadBufferDone: queue the summary, but not for a load that failed."""
	if _loadSucceeded(buffer, args, kwargs):
		_scheduleSummary(buffer)


def _summaryAfterDocumentLoad(buffer, args, kwargs):
	"""event_documentLoadComplete: the second chance for a buffer that loaded empty."""
	_scheduleSummary(buffer)


def _isEnteringDocument(self, args, kwargs):
	"""True when a browse mode focus event is focus arriving from outside the document.

	NVDA sets ``_enteringFromOutside`` in ``event_focusEntered`` when focus passes
	through the document root, and clears it in the first line of ``event_gainFocus``,
	so it has to be read before the original runs. This is the one focus event per
	visit that NVDA answers by reading the line the focus landed on, and it is what
	reads the first line of an Outlook message when the message is opened. Every other
	focus event in the document is ordinary tab and arrow navigation, which must keep
	speaking.

	Deliberately not keyed off ``_hadFirstGainFocus`` as well. That flag is only
	cleared in ``event_treeInterceptor_gainFocus``, so a document whose tree
	interceptor never sees that event would keep it set for good and every focus event
	in the document would be silenced, which is far worse than one line too many.
	``_enteringFromOutside`` is cleared by NVDA on every single call, so it cannot
	latch that way.
	"""
	try:
		return bool(getattr(self, "_enteringFromOutside", False))
	except Exception:
		return False


def _isADifferentDocument(obj):
	"""Whether C{obj} is a document other than the one the user is already reading.

	Escaping out of the browser's find bar puts the focus back in the page, and NVDA
	names the page on the way in. That is the title this add-on exists to silence, and
	it is not news: it is where the user already was. An *embedded* document is news,
	because the keyboard has gone somewhere they did not put it.

	The two are told apart by asking the tree interceptor what document it is showing.
	In Chromium an iframe shares one virtual buffer with the page that contains it, so
	the page is the buffer's ``rootNVDAObject`` and the iframe is not.

	Anything unclear answers True, so the announcement is made. Saying a title one time
	too many is a small annoyance; not saying it is what left the user with no way to
	tell why control+F had stopped opening the find bar.
	"""
	try:
		treeInterceptor = api.getFocusObject().treeInterceptor
	except Exception:
		return True
	root = getattr(treeInterceptor, "rootNVDAObject", None)
	if root is None:
		return True
	try:
		return obj != root
	except Exception:
		return True


def _documentAnnouncementIsExpected():
	"""Whether a document announcing itself right now is one of the moments we silence.

	True while a hooked document call is on the stack, while a buffer is loading or has
	just finished, and while a new foreground window is being announced. Those are the
	three occasions this add-on exists for, and between them they cover every document
	announcement it set out to remove: the page title on load, on entering a document,
	and on switching to a browser window.

	Outside all three there is nothing going on that would explain a document naming
	itself, so it is not the page the user is already on — it is the focus moving into
	a *different* document, an embedded one, while they were working. That has to be
	heard. An iframe holding the focus swallows the browser's own control+F, so the
	find bar silently refuses to open and the enter afterwards is taken by browse mode
	as "activate what is under the cursor"; with the announcement dropped there was
	nothing at all to say the keyboard had moved.
	"""
	if _inCallDepth > 0:
		return True
	if _gateUntil > 0.0 and time.monotonic() < _gateUntil:
		return True
	return _announcingNow()


def _shouldDropObjectSpeech(args, kwargs):
	"""True for the window, dialog and document titles and toasts we never want spoken.

	Only applies inside Outlook and Chromium based browsers, and only to announcements
	NVDA made on its own initiative.
	"""
	if getMode() == MODE_NORMAL:
		return False
	if _bypassDepth > 0:
		# The add-on speaking for itself. Its own announcements are never the noise it
		# was written to remove, and one of them is a deliberate stand-in for exactly the
		# announcement this drops. See L{GlobalPlugin._sayTheBoxInstead}.
		return False
	obj = kwargs.get("obj", args[0] if args else None)
	if obj is None:
		return False
	if "reason" in kwargs:
		reason = kwargs["reason"]
	elif len(args) > 1:
		reason = args[1]
	else:
		# speakObject defaults to QUERY, which is the user asking rather than NVDA
		# volunteering, so it is never dropped.
		return False
	if reason not in _AUTOMATIC_REASONS:
		return False
	try:
		role = obj.role
	except Exception:
		return False
	if role not in _TITLE_ROLES and role not in _ALERT_ROLES:
		return False
	target = _targetOf(obj)
	if target is None:
		return False
	if target != TARGET_OUTLOOK and role in _WINDOW_TITLE_ROLES and _announcingNow():
		# The title of a window the user has just switched to. Alt+tab between two
		# browser windows is the case that needs this: the switcher names each window
		# as it is cycled through, and the moment one is actually activated NVDA
		# cancels that, so dropping the title afterwards as well left the user with
		# nothing but the fragment the switcher got through. Outlook is the exception,
		# because there the brief description of the focus takes the title's place.
		return False
	if (
		role in _DOCUMENT_TITLE_ROLES
		and not _documentAnnouncementIsExpected()
		and _isADifferentDocument(obj)
	):
		# A document other than the one being read has named itself, while nothing was
		# loading, nothing was being entered and no window had just been switched to.
		# That is the focus landing in an embedded document the user did not go to, and
		# it has to be spoken: it is the only warning that the keyboard has left the
		# page, and that the browser's own find bar will not open until it comes back.
		# Returning to the page itself, such as on escaping out of the find bar, is not
		# this, and stays as silent as it has always been.
		log.debug("Mute Browse Mode: another document took the focus, saying so")
		return False
	return True


### NVDA's find in place of the browser's own
#
# NVDA's find is better than a browser's find bar for reading with: it searches the
# browse mode document from the cursor and leaves the cursor on what it found, so the
# next line is the line after the match, and NVDA+F3 carries on from there. A browser's
# find bar puts the keyboard somewhere else entirely and only scrolls the page.
#
# It is also the one that keeps working. An embedded document that has taken the focus
# swallows the browser's control+F, so the find bar silently refuses to open; NVDA's
# find has no such problem, because it never leaves the buffer.
#
# Only in a web browser, and only where NVDA can actually search: Outlook's message body
# is a browse mode document too, and control+F there is Forward.


def _findIsOursIn(obj):
	"""Whether control+F belongs to NVDA's find in the program C{obj} is part of.

	Microsoft Outlook is ruled out first and on its own, because control+F is Forward
	there and has to arrive as the key the user actually pressed. Neither check box
	reaches into Outlook and neither can be made to.

	The test for it is L{outlookIsCurrent}, which asks what window the user is in, and
	not what application the focused object says it belongs to. Those were the same
	question until Outlook began rendering its own message body in an embedded Edge web
	view: everything inside one belongs to msedgewebview2.exe and is Chromium by every
	test there is, so a message being read looked exactly like a web page, and control+F
	in it opened NVDA's find instead of forwarding the message. See L{_isInOutlookWindow}.

	Then, of what is left, two check boxes, the second a widening of the first:

	* "Control+F opens NVDA's find in a web browser" — browsers only;
	* "Bring up NVDA screen reader find when not in Outlook" — everywhere else. This one
	  wins where both are ticked, being the broader of the two, and everywhere it reaches
	  is somewhere the narrower one already reached.
	"""
	if obj is None:
		return False
	if outlookIsCurrent(obj):
		return False
	if getFindWhenNotOutlook():
		return True
	if getBrowserFind():
		return _isWebBrowser(obj)
	return False


#: Roles that count as an edit box for the purpose of searching one.
_EDIT_ROLES = _members(controlTypes.Role, ("EDITABLETEXT", "DOCUMENT", "TERMINAL"))

_STATE_EDITABLE = getattr(controlTypes.State, "EDITABLE", None)
_STATE_PROTECTED = getattr(controlTypes.State, "PROTECTED", None)


class _TextFieldCursorManager(cursorManager.CursorManager):
	"""Lends NVDA's find to an edit box that has no cursor manager of its own.

	NVDA defines find on ``cursorManager.CursorManager``, and only browse mode plus a
	few app modules mix that in — so in an ordinary edit box there is nothing for
	control+F to invoke, and NVDA+control+F does nothing there either.

	Nothing about the search actually needs a browse mode document, though. NVDA's
	``doFindText`` wants somewhere to make a TextInfo, a ``find`` on that TextInfo, and
	somewhere to put the selection, and every editable object already has all three. So
	this borrows the object and hands it to NVDA's own dialog and its own search: what
	the user gets is the find they already know, not a copy of it.
	"""

	def __init__(self, obj):
		super().__init__()
		self._obj = obj

	def makeTextInfo(self, position):
		return self._obj.makeTextInfo(position)

	def _get_selection(self):
		return self._obj.makeTextInfo(textInfos.POSITION_SELECTION)

	def _set_selection(self, info):
		"""Move the caret to what was found.

		NVDA's own version of this hands ``self`` to braille and vision. Ours must hand
		them the real object instead: this one is a stand-in that was made a moment ago
		for one search and is about to be dropped, and it is not what is on screen.
		"""
		info.updateSelection()
		try:
			import braille
			import review
			import vision

			review.handleCaretMove(info)
			braille.handler.handleCaretMove(self._obj)
			vision.handler.handleCaretMove(self._obj)
		except Exception:
			# The caret has already moved; this is only the reporting of it.
			log.debugWarning("Mute Browse Mode: could not report the find move", exc_info=True)

	def doFindText(self, text, reverse=False, caseSensitive=False, willSayAllResume=False):
		"""Search, then read the line the match is on rather than one line past it.

		NVDA's own version finds the text, puts the caret on it, and then reads a range
		it builds like this::

			info.move(textInfos.UNIT_LINE, 1, endPoint="end")

		— the match, extended forward by one line unit, meaning "the found text and the
		rest of its line". In a virtual buffer that is offset arithmetic and the end
		lands exactly on the start of the next line. A UIA text control is not offset
		arithmetic: ``MoveEndpointByUnit`` normalises an endpoint that is in the middle
		of a unit to the unit boundary *first* and then moves a whole unit, so the end
		overshoots into the following line and NVDA reads on into it. Windows 11 Notepad
		is such a control, and this add-on is the first thing that ever ran NVDA's find
		against one, because until then there was no find in an edit box at all.

		So the search is still entirely NVDA's — this passes ``willSayAllResume=True``,
		which is the one thing in ``doFindText`` that suppresses the speaking and
		nothing else, and leaves finding, moving the caret, cancelling speech, the "not
		found" message and ``_lastFindText`` exactly where they were. Only the range
		that gets spoken is built here, by collapsing to the caret and expanding by a
		line, which means the same thing to every kind of text info.

		Whether anything was found is read back from the caret afterwards rather than
		guessed at: NVDA leaves it alone when there is no match.
		"""
		try:
			before = self.makeTextInfo(textInfos.POSITION_CARET)
		except Exception:
			before = None
		super().doFindText(
			text,
			reverse=reverse,
			caseSensitive=caseSensitive,
			willSayAllResume=True,
		)
		if willSayAllResume:
			# Say all is resuming and will do the reading; NVDA would have stayed quiet
			# here too.
			return
		try:
			after = self.makeTextInfo(textInfos.POSITION_CARET)
			if before is not None and after.compareEndPoints(before, "startToStart") == 0:
				# The caret has not moved, so there was no match and NVDA has said so.
				return
			line = after.copy()
			line.expand(textInfos.UNIT_LINE)
		except Exception:
			log.debugWarning("Mute Browse Mode: could not read the line found", exc_info=True)
			return
		speech.speakTextInfo(line, reason=controlTypes.OutputReason.CARET)


def _findableTextField(obj):
	"""C{obj} itself when it is an edit box NVDA could search, otherwise C{None}.

	Not every text object can be searched. ``TextInfo.find`` is implemented by the
	offset and UIA text infos but the base class only raises ``NotImplementedError``, so
	the ones that cannot are ruled out here — before the user has typed what they were
	looking for, rather than after.

	A password box is left alone. It is an edit box by every other test, but searching
	one is no use to anybody and reading the result aloud even less so.
	"""
	if obj is None:
		return None
	try:
		states = set(obj.states or ())
	except Exception:
		states = set()
	if _hasState(states, _STATE_PROTECTED):
		return None
	try:
		if obj.role not in _EDIT_ROLES and not _hasState(states, _STATE_EDITABLE):
			return None
	except Exception:
		return None
	try:
		info = obj.makeTextInfo(textInfos.POSITION_CARET)
	except Exception:
		return None
	if type(info).find is textInfos.TextInfo.find:
		return None
	return obj


def _activeBrowseModeTreeInterceptor(focus):
	"""The browse mode document interceptor at C{focus}, if there is a live one.

	Ready, and in browse mode rather than passthrough — the same test L{_findSource}
	uses to decide there is a document NVDA's own find could search.

	Used on its own too, for a narrower question: does *any* browse mode document sit
	under the focus, regardless of whether these settings want NVDA's find to open
	here. A browse mode document binds control+F to NVDA's own find by default, Outlook
	included, whether or not this add-on wants that — Outlook's message body is a
	browse mode document like any other, and its default control+F has to be kept from
	ever firing, not merely left unclaimed.
	"""
	treeInterceptor = getattr(focus, "treeInterceptor", None)
	if (
		isinstance(treeInterceptor, cursorManager.CursorManager)
		and getattr(treeInterceptor, "isReady", False)
		and not getattr(treeInterceptor, "passThrough", False)
	):
		return treeInterceptor
	return None


def _findSource():
	"""What control+F would search here, as C{(kind, object)}, or C{None} for nothing.

	C{kind} is ``"document"`` for a browse mode buffer, which NVDA can already search as
	it stands, or ``"field"`` for an edit box, which needs L{_TextFieldCursorManager}
	wrapped round it first. The wrapping is deliberately left to the caller: this runs on
	every focus change to decide whether the key is bound, and there is no point building
	an adapter that is only going to be thrown away.

	Everything has to line up before the key is taken off the program:

	* one of the two find check boxes covers the program this is, and it is not Outlook
	  (L{_findIsOursIn});
	* and there is either a browse mode document in browse mode, or an edit box that can
	  be searched.
	"""
	if not (getBrowserFind() or getFindWhenNotOutlook()):
		return None
	try:
		focus = api.getFocusObject()
	except Exception:
		return None
	# Outlook first and without exception: control+F forwards the message there. The
	# focus object is safe to judge on now, because L{outlookIsCurrent} asks which window
	# it is in rather than which application it says it belongs to.
	if not _findIsOursIn(focus):
		return None
	treeInterceptor = _activeBrowseModeTreeInterceptor(focus)
	if treeInterceptor is not None:
		# Browse mode, so search the whole document rather than the control the focus
		# happens to be sitting on inside it.
		return ("document", treeInterceptor)
	field = _findableTextField(focus)
	return ("field", field) if field is not None else None


def _findTarget():
	"""The object to run NVDA's find on, ready to use, or C{None}."""
	source = _findSource()
	if source is None:
		return None
	kind, obj = source
	return obj if kind == "document" else _TextFieldCursorManager(obj)


### Tracing
#
# Off, and costing nothing, unless a marker file exists beside NVDA's own log. It is
# here because the interesting question about this add-on is never "what did it say"
# but "which of NVDA's handlers ran, and on what", and NVDA's log cannot answer that
# when the user has logging turned off, which is the normal setting.
#
# Everything below is wrapped so that tracing can never be the thing that breaks a
# keystroke: any failure switches tracing off rather than propagating.

_TRACE_MARKER = os.path.join(tempfile.gettempdir(), "muteBrowseModeTrace.on")
_TRACE_PATH = os.path.join(tempfile.gettempdir(), "muteBrowseModeTrace.log")

try:
	_tracing = os.path.exists(_TRACE_MARKER)
except Exception:
	_tracing = False


def _traceWrite(line):
	global _tracing
	if not _tracing:
		return
	try:
		with open(_TRACE_PATH, "a", encoding="utf-8", errors="replace") as f:
			f.write("%s %s\n" % (time.strftime("%H:%M:%S"), line))
	except Exception:
		# A trace that cannot be written must not keep trying on every keystroke.
		_tracing = False


def _rootWindowClassOf(obj):
	"""The class of the top level window C{obj} is in, for the trace to show.

	Two user32 calls and nothing else, so it is safe to ask from the input hook thread,
	which is where gestures are traced from. It is worth a field of its own because the
	application an object belongs to is no longer the application it is *in* — an Outlook
	message body rendered in a web view says msedgewebview2, and only the window it sits
	in says Outlook.
	"""
	try:
		import winUser

		root = _rootWindowOf(obj)
		return winUser.getClassName(root) if root else ""
	except Exception:
		return "?"


def _describe(obj):
	"""Name, role, window classes and application of C{obj}, each one guarded."""
	parts = []
	for label, getter in (
		("name", lambda: repr((obj.name or "")[:40])),
		("role", lambda: str(obj.role)),
		("class", lambda: _windowClassOf(obj)),
		("app", lambda: _appNameOf(obj)),
		("rootClass", lambda: _rootWindowClassOf(obj)),
		("states", lambda: ",".join(sorted(str(s) for s in (obj.states or ())))[:80]),
	):
		try:
			parts.append("%s=%s" % (label, getter()))
		except Exception:
			parts.append("%s=?" % label)
	return " ".join(parts)


def _describeScript(func):
	if func is None:
		return "script=None(passes to the program)"
	try:
		owner = getattr(func, "__self__", None)
		return "script=%s.%s on %s" % (
			getattr(func, "__module__", "?"),
			getattr(func, "__qualname__", getattr(func, "__name__", "?")),
			type(owner).__name__ if owner is not None else "-",
		)
	except Exception:
		return "script=?"


def _traceSpeech(args, kwargs):
	"""Record what is about to be spoken, and by whom.

	Every route into the synthesiser goes through ``speech.speak``, so this catches the
	lot: NVDA's own announcements, the add-on's, and anything an app module adds. The
	caller is worth as much as the text — "which line was read" and "who read it" are
	different questions, and a find that reads the wrong line looks identical from the
	outside to a find that reads the right one and is then talked over by something else.
	"""
	if not _tracing:
		return
	try:
		sequence = args[0] if args else kwargs.get("speechSequence")
		if not isinstance(sequence, (list, tuple)):
			return
		text = " ".join(part for part in sequence if isinstance(part, str)).strip()
		if not text:
			return
		caller = "?"
		try:
			import inspect

			frame = inspect.currentframe().f_back.f_back
			for _step in range(6):
				if frame is None:
					break
				name = frame.f_code.co_name
				if name not in ("speak", "_traceSpeech"):
					caller = "%s:%s" % (frame.f_globals.get("__name__", "?"), name)
					break
				frame = frame.f_back
		except Exception:
			pass
		_traceWrite("SAY [%s] %r" % (caller, text[:160]))
	except Exception:
		pass


def _traceGesture(gesture):
	"""One line saying what NVDA is about to do with this key, and to what."""
	if not _tracing:
		return
	try:
		identifier = "?"
		try:
			identifier = gesture.identifiers[0]
		except Exception:
			pass
		try:
			script = scriptHandler.findScript(gesture)
		except Exception:
			script = None
		focus = None
		try:
			focus = api.getFocusObject()
		except Exception:
			pass
		ti = getattr(focus, "treeInterceptor", None)
		_traceWrite(
			"KEY %s | %s | focus: %s | ti=%s passThrough=%s isReady=%s"
			% (
				identifier,
				_describeScript(script),
				_describe(focus) if focus is not None else "?",
				type(ti).__name__ if ti is not None else "None",
				getattr(ti, "passThrough", "-"),
				getattr(ti, "isReady", "-"),
			),
		)
	except Exception:
		pass


def _onGesture(*args, **kwargs):
	"""Any key the user presses means they want to hear things again.

	Registered with the decide_executeGesture extension point purely to get told
	about input; it never vetoes a gesture.
	"""
	_traceGesture(kwargs.get("gesture", args[0] if args else None))
	if _gateUntil > 0.0:
		_closeGate()
	# The user is driving again, so there is no window announcement left to wait for.
	# This also stops a cancelled announcement holding the gate down to its ceiling.
	_endAnnouncement()
	# A page summary that has not been spoken yet is now stale: the user has stopped
	# waiting for the page and started using it.
	_cancelSummary()
	return True


### Settings

def _selectionForMode(mode):
	try:
		return MODES.index(mode)
	except ValueError:
		return MODES.index(MODE_NORMAL)


def _selectionForSummaryMode(mode):
	try:
		return SUMMARY_MODES.index(mode)
	except ValueError:
		return SUMMARY_MODES.index(SUMMARY_SPEAK)


def _addChoices(panel, sHelper):
	"""Add both of this add-on's combo boxes to C{panel}."""
	panel._muteBrowseModeChoice = sHelper.addLabeledControl(
		# Translators: Label of a combo box added to NVDA's Browse Mode settings.
		_("Mute &browse mode:"),
		wx.Choice,
		choices=getModeLabels(),
	)
	panel._muteBrowseModeChoice.SetSelection(_selectionForMode(getMode()))
	panel._muteBrowseModeSummaryChoice = sHelper.addLabeledControl(
		# Translators: Label of a combo box added to NVDA's Browse Mode settings.
		_("Announce a &page summary when a page has loaded:"),
		wx.Choice,
		choices=getSummaryModeLabels(),
	)
	panel._muteBrowseModeSummaryChoice.SetSelection(_selectionForSummaryMode(getSummaryMode()))
	panel._muteBrowseModeLoadingCheckBox = sHelper.addItem(
		wx.CheckBox(
			panel,
			# Translators: Label of a check box added to NVDA's Browse Mode settings.
			label=_('Say "&loading complete" before the page summary'),
		),
	)
	panel._muteBrowseModeLoadingCheckBox.SetValue(getAnnounceLoadingComplete())
	panel._muteBrowseModeLinksCheckBox = sHelper.addItem(
		wx.CheckBox(
			panel,
			# Translators: Label of a check box added to NVDA's Browse Mode settings.
			label=_("Links are on their &own line"),
		),
	)
	panel._muteBrowseModeLinksCheckBox.SetValue(getLinksOnOwnLine())
	panel._muteBrowseModeFindCheckBox = sHelper.addItem(
		wx.CheckBox(
			panel,
			# Translators: Label of a check box added to NVDA's Browse Mode settings.
			label=_("Control+F opens NVDA's &find in a web browser"),
		),
	)
	panel._muteBrowseModeFindCheckBox.SetValue(getBrowserFind())
	panel._muteBrowseModeFindAnywhereCheckBox = sHelper.addItem(
		wx.CheckBox(
			panel,
			# Translators: Label of a check box added to NVDA's Browse Mode settings.
			label=_("Bring up NVDA screen reader find when &not in Outlook"),
		),
	)
	panel._muteBrowseModeFindAnywhereCheckBox.SetValue(getFindWhenNotOutlook())


def _saveChoices(panel):
	choice = getattr(panel, "_muteBrowseModeChoice", None)
	if choice is not None:
		index = choice.GetSelection()
		if 0 <= index < len(MODES):
			setMode(MODES[index])
	choice = getattr(panel, "_muteBrowseModeSummaryChoice", None)
	if choice is not None:
		index = choice.GetSelection()
		if 0 <= index < len(SUMMARY_MODES):
			setSummaryMode(SUMMARY_MODES[index])
	checkBox = getattr(panel, "_muteBrowseModeLoadingCheckBox", None)
	if checkBox is not None:
		setAnnounceLoadingComplete(checkBox.IsChecked())
	checkBox = getattr(panel, "_muteBrowseModeLinksCheckBox", None)
	if checkBox is not None:
		setLinksOnOwnLine(checkBox.IsChecked())
	checkBox = getattr(panel, "_muteBrowseModeFindCheckBox", None)
	if checkBox is not None:
		setBrowserFind(checkBox.IsChecked())
	checkBox = getattr(panel, "_muteBrowseModeFindAnywhereCheckBox", None)
	if checkBox is not None:
		setFindWhenNotOutlook(checkBox.IsChecked())


def _refreshChoices(panel):
	choice = getattr(panel, "_muteBrowseModeChoice", None)
	if choice is not None:
		choice.SetSelection(_selectionForMode(getMode()))
	choice = getattr(panel, "_muteBrowseModeSummaryChoice", None)
	if choice is not None:
		choice.SetSelection(_selectionForSummaryMode(getSummaryMode()))
	checkBox = getattr(panel, "_muteBrowseModeLoadingCheckBox", None)
	if checkBox is not None:
		checkBox.SetValue(getAnnounceLoadingComplete())
	checkBox = getattr(panel, "_muteBrowseModeLinksCheckBox", None)
	if checkBox is not None:
		checkBox.SetValue(getLinksOnOwnLine())
	checkBox = getattr(panel, "_muteBrowseModeFindCheckBox", None)
	if checkBox is not None:
		checkBox.SetValue(getBrowserFind())
	checkBox = getattr(panel, "_muteBrowseModeFindAnywhereCheckBox", None)
	if checkBox is not None:
		checkBox.SetValue(getFindWhenNotOutlook())


def _makeSettingsWrapper(original):
	def makeSettings(self, settingsSizer):
		original(self, settingsSizer)
		try:
			_addChoices(self, guiHelper.BoxSizerHelper(self, sizer=settingsSizer))
		except Exception:
			self._muteBrowseModeChoice = None
			self._muteBrowseModeSummaryChoice = None
			self._muteBrowseModeLoadingCheckBox = None
			self._muteBrowseModeLinksCheckBox = None
			self._muteBrowseModeFindCheckBox = None
			self._muteBrowseModeFindAnywhereCheckBox = None
			log.error(
				"Mute Browse Mode: could not add the controls to Browse Mode settings",
				exc_info=True,
			)

	return makeSettings


def _onSaveWrapper(original):
	def onSave(self):
		original(self)
		_saveChoices(self)

	return onSave


def _onPanelActivatedWrapper(original):
	def onPanelActivated(self):
		# NVDA reuses the settings dialog, so refresh in case a setting was changed
		# elsewhere (a profile switch, or one of the cycle scripts) since it was built.
		_refreshChoices(self)
		original(self)

	return onPanelActivated


def _patchBrowseModePanel():
	panel = settingsDialogs.BrowseModePanel
	_patch(panel, "makeSettings", _makeSettingsWrapper(panel.makeSettings))
	_patch(panel, "onSave", _onSaveWrapper(panel.onSave))
	_patch(panel, "onPanelActivated", _onPanelActivatedWrapper(panel.onPanelActivated))


class MuteBrowseModePanel(settingsDialogs.SettingsPanel):
	"""Only registered if the combo boxes could not be injected into Browse Mode."""

	# Translators: Title of the add-on's own settings category, used as a fallback.
	title = _("Mute Browse Mode")

	def makeSettings(self, settingsSizer):
		_addChoices(self, guiHelper.BoxSizerHelper(self, sizer=settingsSizer))

	def onSave(self):
		_saveChoices(self)

	def onPanelActivated(self):
		_refreshChoices(self)
		super().onPanelActivated()


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	# Translators: Category for this add-on's commands in the Input Gestures dialog.
	scriptCategory = _("Mute Browse Mode")

	def __init__(self):
		super().__init__()
		self._panelPatched = False
		self._ownPanelAdded = False
		self._gestureHandlerRegistered = False
		#: Whether the window we switched away from was an Outlook one.
		self._lastForegroundWasOutlook = False
		#: Latest moment the focus landing in Outlook still counts as having just
		#: arrived there. A deadline rather than a flag, so it cannot latch.
		self._outlookArrivalUntil = 0.0
		#: The window the message body announcement was last made for, so that tabbing
		#: away and back announces it again but a repeated focus event does not.
		self._lastBodyWindow = None
		#: The window and word the spelling checker was last answered for, so that the
		#: dialog moving on to the next word says that one and a repeated focus event
		#: does not say the same one twice.
		self._lastSpellCheck = None
		#: Whether the focus was inside the spelling dialog last time we looked. A check
		#: that starts while the last one is still remembered would otherwise have its
		#: first word deduplicated away against it.
		self._inSpellDialog = False
		#: The window the end of a check was last announced for, so that a dialog raising
		#: a second focus event for the same arrival does not say it twice. Forgotten as
		#: soon as the focus goes anywhere else. See L{_finishSpellCheck}.
		self._saidCompleteFor = None
		#: After spell-check completion, suppress the automatic message-body announcement
		#: that Outlook emits when focus returns to the message body.
		self._suppressBodyUntil = 0.0
		#: Short window for suppressing NVDA's separate automatic OK-button speech.
		self._suppressSpellOkUntil = 0.0
		#: Whether control+F is currently bound to this plugin. Only true where there is
		#: a browse mode document NVDA's find could search; see
		#: L{_syncBrowserFindBinding}.
		self._browserFindBound = False

		try:
			# Gating speech itself, rather than each announcement, catches every route
			# into the synthesiser: speakTextInfo, speakObject, speakMessage and
			# ui.message all funnel through speech.speech.speak.
			speakWrapper = self._makeSpeakWrapper(speech.speech.speak)
			_patch(speech.speech, "speak", speakWrapper)
			_patch(speech, "speak", speakWrapper)

			# Window titles, document titles and toasts in Outlook and Chromium.
			speakObjectWrapper = self._makeSpeakObjectWrapper(speech.speech.speakObject)
			_patch(speech.speech, "speakObject", speakObjectWrapper)
			_patch(speech, "speakObject", speakObjectWrapper)

			# The document announcement must not cut off the window title in front of it.
			cancelWrapper = self._makeCancelSpeechWrapper(speech.speech.cancelSpeech)
			_patch(speech.speech, "cancelSpeech", cancelWrapper)
			_patch(speech, "cancelSpeech", cancelWrapper)
		except Exception:
			log.error("Mute Browse Mode: could not install speech hooks", exc_info=True)
			_unpatchAll()
			return

		# Each of the remaining hooks is optional and independent: if one of them does
		# not fit this NVDA, the rest still do their job.
		self._installDocumentHooks()
		self._installProgressHook()
		self._installNoiseHooks()

		try:
			inputCore.decide_executeGesture.register(_onGesture)
			self._gestureHandlerRegistered = True
		except Exception:
			log.error("Mute Browse Mode: could not hook input gestures", exc_info=True)

		# NVDA may well be starting with a browser already in front, in which case no
		# foreground change is coming to tell us about it.
		self._syncBrowserFindBinding()

		try:
			_patchBrowseModePanel()
			self._panelPatched = True
		except Exception:
			log.error(
				"Mute Browse Mode: could not extend the Browse Mode settings panel, "
				"falling back to a separate settings category",
				exc_info=True,
			)
			try:
				settingsDialogs.NVDASettingsDialog.categoryClasses.append(MuteBrowseModePanel)
				self._ownPanelAdded = True
			except Exception:
				log.error("Mute Browse Mode: could not add a settings panel at all", exc_info=True)

	def _installDocumentHooks(self):
		"""The four moments NVDA announces a browse mode document."""
		hooks = (
			# Focus arriving in a browse mode document from outside it. NVDA answers
			# this by speaking the line the focus landed on, which is what reads the
			# first line of an Outlook message. It is a focus event rather than part of
			# a load, so none of the load hooks below can catch it.
			(
				browseMode.BrowseModeDocumentTreeInterceptor,
				"event_gainFocus",
				dict(
					after=_TRAILING_GATE,
					chime=True,
					onlyIf=_isEnteringDocument,
				),
			),
			# Entering a browse mode document for the first time: web pages, and
			# Outlook and Word message bodies, which are browse mode documents but not
			# virtual buffers. This is where the name, the word "document" and the
			# first line of the buffer are spoken.
			(
				browseMode.BrowseModeDocumentTreeInterceptor,
				"event_treeInterceptor_gainFocus",
				dict(after=_TRAILING_GATE, chime=True),
			),
			# A virtual buffer starting to load, which is what silences "Loading
			# document...", and the point the page summary is owed from.
			(
				virtualBuffers.VirtualBuffer,
				"loadBuffer",
				dict(after=_LOAD_GATE, post=_armSummary),
			),
			# ...and finishing, which speaks "Refreshed" on a reload and is the moment
			# both the chime and the page summary belong to.
			(
				virtualBuffers.VirtualBuffer,
				"_loadBufferDone",
				dict(
					after=_TRAILING_GATE,
					chime=True,
					chimeCheck=_loadSucceeded,
					post=_summaryAfterLoad,
				),
			),
		)
		for owner, name, options in hooks:
			try:
				_hookGate(owner, name, **options)
			except Exception:
				log.error("Mute Browse Mode: could not hook %s.%s" % (owner.__name__, name), exc_info=True)

		# A buffer that finished loading empty is one NVDA is still waiting on; it
		# reports the document from this event instead, and so do we.
		try:
			_hookAfter(virtualBuffers.VirtualBuffer, "event_documentLoadComplete", _summaryAfterDocumentLoad)
		except Exception:
			log.error("Mute Browse Mode: could not hook event_documentLoadComplete", exc_info=True)

		# Where a line starts and ends, so that an Outlook message rendered as a web
		# document can put each of its links and buttons on a line of its own.
		try:
			bufferText = virtualBuffers.VirtualBufferTextInfo
			_patch(bufferText, "_getLineOffsets", _makeLineOffsetsWrapper(bufferText._getLineOffsets))
		except Exception:
			log.error("Mute Browse Mode: could not hook _getLineOffsets", exc_info=True)

		# ...and the arrow keys, for an Outlook message rendered by Word, which has no
		# buffer and so no line offsets to work out.
		for name, direction in (
			("script_moveByLine_forward", 1),
			("script_moveByLine_back", -1),
		):
			try:
				manager = cursorManager.CursorManager
				_patch(manager, name, _makeMoveByLineWrapper(getattr(manager, name), direction))
			except Exception:
				log.error("Mute Browse Mode: could not hook %s" % name, exc_info=True)

	def _installProgressHook(self):
		"""Let a progress bar past the gate. See the notes above L{_makeProgressBarWrapper}."""
		try:
			from NVDAObjects import behaviors

			_patch(
				behaviors.ProgressBar,
				"event_valueChange",
				_makeProgressBarWrapper(behaviors.ProgressBar.event_valueChange),
			)
		except Exception:
			log.error("Mute Browse Mode: could not hook the progress bar", exc_info=True)

	def _installNoiseHooks(self):
		"""Live regions and alerts, silenced in Outlook and Chromium only.

		Web pages announce "flash" messages through ARIA live regions and alerts, and
		browsers use the same machinery for their own toasts. These are the calls that
		speak them; they are left to run so braille still shows the text.
		"""
		targets = []
		try:
			import NVDAObjects

			targets.append((NVDAObjects.NVDAObject, "event_liveRegionChange"))
		except Exception:
			log.error("Mute Browse Mode: could not reach NVDAObject", exc_info=True)
		try:
			from NVDAObjects import behaviors

			# event_show is a second class attribute pointing at event_alert, so
			# patching event_alert alone would leave it live.
			targets.append((behaviors.Notification, "event_alert"))
			targets.append((behaviors.Notification, "event_show"))
		except Exception:
			log.error("Mute Browse Mode: could not reach the Notification behaviour", exc_info=True)
		try:
			from NVDAObjects.IAccessible import IAccessible

			targets.append((IAccessible, "event_alert"))
		except Exception:
			log.error("Mute Browse Mode: could not reach IAccessible", exc_info=True)

		for owner, name in targets:
			try:
				_hookSilent(owner, name)
			except Exception:
				log.error("Mute Browse Mode: could not hook %s.%s" % (owner.__name__, name), exc_info=True)

	def event_foreground(self, obj, nextHandler):
		"""A different window has come to the front: alt+tab, or anything like it.

		Two things happen here. The deadline gate stands down until NVDA has finished
		announcing the new window, so the title and the control the focus landed on are
		never cut off. And if this is Outlook arriving from somewhere else, the focus
		event that follows is answered with a brief description of where the user has
		landed, in place of NVDA's own report.
		"""
		self._syncBrowserFindBinding()
		# A window has changed, so the message body the user was last told about is no
		# longer the one they are in. Coming back to a message announces it again, which
		# matters now that NVDA's own announcement of the body is dropped: without this,
		# switching away and back would land in the body in silence.
		self._lastBodyWindow = None
		try:
			isOutlook = _isInOutlookWindow(obj)
			if getMode() != MODE_NORMAL:
				# In normal mode nothing is being silenced, so there is nothing to
				# stand down and no reason to touch NVDA's speech at all.
				_beginAnnouncement()
			if isOutlook and not self._lastForegroundWasOutlook and getMode() != MODE_NORMAL:
				self._outlookArrivalUntil = time.monotonic() + _OUTLOOK_ARRIVAL_WINDOW
				log.debug("Mute Browse Mode: Outlook is coming to the front")
			self._lastForegroundWasOutlook = isOutlook
		except Exception:
			log.error("Mute Browse Mode: could not handle the foreground change", exc_info=True)
		nextHandler()

	def event_gainFocus(self, obj, nextHandler):
		"""Describe where the focus landed, once Outlook has fully come to the front.

		NVDA's own report still runs, silently, so its property caches and braille are
		exactly as they would have been; only the speech is replaced. If the brief
		report cannot be built, nothing is replaced and NVDA speaks as usual.
		"""
		self._syncBrowserFindBinding()
		if _tracing:
			ti = getattr(obj, "treeInterceptor", None)
			_traceWrite(
				"FOCUS %s | ti=%s passThrough=%s"
				% (
					_describe(obj),
					type(ti).__name__ if ti is not None else "None",
					getattr(ti, "passThrough", "-"),
				),
			)
		# The dialog saying the check has finished, which is asked about before the
		# checker itself: it comes up whether or not anything was found, so a check that
		# found nothing at all arrives here with no word ever having been said.
		if self._isSpellCheckComplete(obj):
			with _hardMute():
				nextHandler()
			self._finishSpellCheck(obj)
			return
		# The focus is somewhere other than that dialog, so the next time it lands on one
		# it is a new arrival and worth speaking for. See L{_finishSpellCheck}.
		self._saidCompleteFor = None

		# The F7 spelling dialog. NVDA's own report of the box holding the word says the
		# label, the word and the spelling at one speed, so it is answered here, before
		# nextHandler, rather than added to afterwards. Everything else in the dialog —
		# the suggestions, the buttons — is left to NVDA to report as usual, so tabbing
		# around it works normally.
		#
		# The box raises a focus event every time its word changes rather than a value
		# change, which is what carries the dialog from one mistake to the next: each new
		# word arrives here as if the user had just landed on the box.
		spellField = _outlookSpellCheckField(obj)
		if spellField is not None:
			if not self._inSpellDialog:
				# A check that has only just started. What the last one finished on must
				# not be allowed to count as "already said" against the first word of
				# this one, or opening the dialog on the same word twice running says it
				# once.
				self._inSpellDialog = True
				self._lastSpellCheck = None
			word = _misspelledWord(spellField)
			seen = (getattr(spellField, "windowHandle", None), word) if word is not None else None
			isNewWord = seen is not None and seen != self._lastSpellCheck
			onTheBox = _windowClassOf(obj) in _WORD_DIALOG_WINDOW_CLASSES
			if isNewWord or onTheBox:
				# Silencing NVDA is only ever safe where there is something to say
				# instead. On the box there is: if no word can be worked out at all, the
				# looks below end by reading the box out the way NVDA would have.
				with _hardMute():
					nextHandler()
			else:
				nextHandler()
			if isNewWord:
				self._lastSpellCheck = seen
				log.debug("Mute Browse Mode: spelling checker is asking about %r" % word)
				_announceMisspelledWord(word)
			elif word is None and onTheBox:
				# The dialog is on screen before Word has picked the error out in the
				# message, so at the moment it takes the focus there is nothing to read
				# yet. That is what left the first word of a check announced as nothing
				# at all. Look again in a moment rather than giving up on it.
				self._lookAgainForTheWord(getattr(spellField, "windowHandle", None), 0)
			return
		self._inSpellDialog = False

		# If Outlook returns focus directly to the message body, silence NVDA's
		# document/body speech BEFORE nextHandler() runs. Otherwise the document text
		# has already been spoken by the time _reportMessageBody gets a chance to act.
		if self._lastSpellCheck is not None and _isOutlookMessageBody(obj):
			with _hardMute():
				nextHandler()
			# No dialog and no button: Outlook has simply handed the message back.
			self._finishSpellCheck(onTheOkButton=False)
			# The body is not the dialog, so nothing is being guarded against a repeat.
			self._saidCompleteFor = None
			return

		# Tabbing into the body of a message. NVDA announces the Word editing surface
		# itself first — "document, page 1, section 1, blank" — and the add-on's own
		# "you are now in the message body" only followed it, so the announcement this
		# add-on exists to remove was still the first thing heard. Decided before
		# nextHandler(), the same way the spelling dialog above is, so that it never
		# reaches the synthesiser at all. Said afterwards, because the browse mode
		# document a message body can be cancels speech on its way into focus mode and
		# would cut off anything said first.
		if getMode() != MODE_NORMAL and _isOutlookMessageBody(obj):
			with _hardMute():
				nextHandler()
			# Outlook has settled on a real control, so there is no arrival left to
			# describe: the body is where the user has landed, and it is about to say so.
			self._outlookArrivalUntil = 0.0
			self._reportMessageBody(obj)
			return

		if time.monotonic() < self._outlookArrivalUntil and _isInOutlookWindow(obj):
			try:
				role = getattr(obj, "role", None)
				if role in _TRANSIENT_FOCUS_ROLES:
					# Outlook has not settled on a real control yet. Keep waiting.
					nextHandler()
					return
				sequence = _briefFocusSpeech(obj)
			except Exception:
				log.error("Mute Browse Mode: could not describe the Outlook focus", exc_info=True)
				sequence = None
			self._outlookArrivalUntil = 0.0
			if sequence:
				with _hardMute():
					nextHandler()
				_speakBriefFocus(sequence)
				self._reportMessageBody(obj)
				return
		nextHandler()
		self._reportMessageBody(obj)

	def _lookAgainForTheWord(self, window, attempt):
		"""Ask the spelling dialog again for the word it is asking about.

		The dialog takes the focus before Word has selected the error in the message, so
		the word cannot always be read at the moment it arrives. Each look is scheduled
		rather than waited for, so nothing is ever held up, and they stop as soon as the
		focus leaves the dialog.

		The box is remembered as a window rather than as an object, because NVDA keeps an
		object's answers once it has given them, and the whole point of asking again is to
		get an answer that has changed since.
		"""
		if not window or attempt >= len(_SPELL_RETRY_DELAYS):
			return
		core.callLater(_SPELL_RETRY_DELAYS[attempt], self._retryTheWord, window, attempt)

	def _retryTheWord(self, window, attempt):
		"""One of those later looks. See L{_lookAgainForTheWord}."""
		try:
			if not self._stillInSpellDialog(window):
				return
			field = _objectFromWindow(window)
			word = _misspelledWord(field) if field is not None else None
			if word is None:
				if attempt + 1 < len(_SPELL_RETRY_DELAYS):
					self._lookAgainForTheWord(window, attempt + 1)
				else:
					self._sayTheBoxInstead()
				return
			seen = (window, word)
			if seen == self._lastSpellCheck:
				return
			self._lastSpellCheck = seen
			log.debug("Mute Browse Mode: spelling checker is asking about %r" % word)
			_announceMisspelledWord(word)
		except Exception:
			log.error("Mute Browse Mode: could not look again for the word", exc_info=True)

	def _stillInSpellDialog(self, window):
		"""Whether the focus is still in the spelling dialog the box C{window} is in.

		A later look must never speak into a window the user has already moved on to.
		"""
		try:
			focus = api.getFocusObject()
		except Exception:
			return False
		current = _outlookSpellCheckField(focus)
		if current is None:
			return False
		return getattr(current, "windowHandle", None) == window

	def _sayTheBoxInstead(self):
		"""Say what NVDA would have said, when no word could be worked out at all.

		Only reached where the box holding the word was silenced and none of the later
		looks found anything in it. Reading out the sentence the error sits in is not
		what this add-on is for, but it is what NVDA has always done, and it is a great
		deal better than the dialog arriving in silence.
		"""
		try:
			focus = api.getFocusObject()
		except Exception:
			return
		if _windowClassOf(focus) not in _WORD_DIALOG_WINDOW_CLASSES:
			# Something else in the dialog has the focus, and NVDA has already reported
			# it in the ordinary way.
			return
		log.debug("Mute Browse Mode: no word in the spelling dialog, saying the box")
		with _ownSpeech():
			speech.speakObject(focus, reason=controlTypes.OutputReason.FOCUS)

	def _isSpellCheckOk(self, obj):
		"""Whether C{obj} is the OK button of the dialog the words were asked in.

		Only an answer where a check was being followed, and only where the button
		belongs to the same window the last word was asked about. Both of those are why
		it cannot answer for a check that found nothing at all.
		"""
		if self._lastSpellCheck is None:
			return False
		try:
			lastWindow = self._lastSpellCheck[0]
			window = getattr(obj, "windowHandle", None)
			return (
				window == lastWindow
				or winUser.getAncestor(window, winUser.GA_ROOT)
				== winUser.getAncestor(lastWindow, winUser.GA_ROOT)
			)
		except Exception:
			return False

	def _isSpellCheckComplete(self, obj):
		"""Whether C{obj} is the OK button of the dialog saying the check has finished.

		Asked two ways, because there are two ways of arriving at this dialog. A check
		that found something ends where the words were being asked, so the button belongs
		to a window the add-on already knows about. A check that found nothing at all
		never said a word, so there is nothing to match it against and the only thing that
		can identify it is what it says — which is also the surer of the two, and the
		reason the dialog is read rather than assumed.
		"""
		if not _isOkButton(obj):
			return False
		if self._isSpellCheckOk(obj):
			return True
		return _saysSpellCheckIsComplete(obj)

	def _finishSpellCheck(self, obj=None, onTheOkButton=True):
		"""Say the check is over, once, and stop everything that was following it.

		NVDA announces the OK button a second time of its own accord, and Outlook returns
		the focus to the message afterwards, which would otherwise be answered with "you
		are now in the message body". Both are held off for a moment rather than switched
		off, so nothing can be left silenced.

		Once only. That dialog raises more than one focus event for the same arrival, and
		each of them is the end of the same check, so the window it was said for is
		remembered. The memory is of a window rather than of a moment in time, and it is
		forgotten the instant the focus goes anywhere else, so running a second check
		straight after the first always says so again.
		"""
		window = getattr(obj, "windowHandle", None) if obj is not None else None
		alreadySaid = window is not None and window == self._saidCompleteFor
		self._saidCompleteFor = window
		self._lastSpellCheck = None
		self._inSpellDialog = False
		self._suppressBodyUntil = time.monotonic() + 2.0
		self._suppressSpellOkUntil = time.monotonic() + 2.0
		if alreadySaid:
			if _tracing:
				_traceWrite("SPELLDONE already said for window %s" % window)
			return
		log.debug("Mute Browse Mode: the spelling check has finished")
		with _ownSpeech():
			speech.speak(_spellCheckCompleteSpeech(onTheOkButton))

	def _reportMessageBody(self, obj):
		"""Say when the focus has reached an Outlook message body.

		Tabbing through a new message goes To, Cc, Subject, body, and NVDA names the
		first three but gives the body no name to read, so there is nothing to tell the
		user they have arrived in the part they are meant to type into.

		Said after NVDA's own announcement rather than before it, because the browse
		mode document a message body can be will cancel speech on its way into focus
		mode, and anything said first would be cut off by that.
		"""
		if getMode() == MODE_NORMAL:
			return
		try:
			if not _isOutlookMessageBody(obj):
				self._lastBodyWindow = None
				return
			if time.monotonic() < self._suppressBodyUntil:
				self._lastBodyWindow = getattr(obj, "windowHandle", None)
				return
			window = getattr(obj, "windowHandle", None)
			if window is not None and window == self._lastBodyWindow:
				# The same body raising a second focus event, not a new arrival.
				return
			self._lastBodyWindow = window
			_announceMessageBody()
		except Exception:
			log.error("Mute Browse Mode: could not announce the message body", exc_info=True)

	def terminate(self):
		_cancelSummary()
		_resetGates()
		if self._ownPanelAdded:
			try:
				settingsDialogs.NVDASettingsDialog.categoryClasses.remove(MuteBrowseModePanel)
			except Exception:
				log.error("Mute Browse Mode: could not remove the settings panel", exc_info=True)
		if self._gestureHandlerRegistered:
			try:
				inputCore.decide_executeGesture.unregister(_onGesture)
			except Exception:
				log.error("Mute Browse Mode: could not unhook input gestures", exc_info=True)
		_unpatchAll()
		super().terminate()

	def _spellCheckIsRunning(self):
		"""Whether the Outlook spelling dialog is part way through a check we are following."""
		return self._lastSpellCheck is not None or time.monotonic() < self._suppressSpellOkUntil

	def _makeSpeakWrapper(self, original):
		def speak(*args, **kwargs):
			muting = getMode() != MODE_NORMAL
			if muting and _isGated():
				return
			if muting and self._spellCheckIsRunning() and _isBareUnavailable(args, kwargs):
				# The spelling dialog disabling itself on its way to saying it has
				# finished. See L{_isBareUnavailable}.
				return
			if _tracing:
				_traceSpeech(args, kwargs)
			if muting:
				try:
					args, kwargs = _filterColumnLabels(args, kwargs)
				except Exception:
					log.debugWarning(
						"Mute Browse Mode: could not filter the column headings",
						exc_info=True,
					)
			if _announcingNow():
				# Let the announcement tell us when it has actually been said, rather
				# than guessing how long a window title takes to read out.
				try:
					args, kwargs = _tagAnnouncement(args, kwargs)
				except Exception:
					log.debugWarning("Mute Browse Mode: could not tag the announcement", exc_info=True)
			return original(*args, **kwargs)

		speak.__name__ = "speak"
		speak.__doc__ = getattr(original, "__doc__", None)
		return speak

	def _makeCancelSpeechWrapper(self, original):
		"""Stop the document announcement cancelling the window title before it.

		``BrowseModeDocumentTreeInterceptor.event_gainFocus`` cancels speech when the
		focus lands somewhere that puts the document into focus mode, on the reasoning
		that a focus change should stop the page being read aloud. Arriving from
		another window there is no page being read aloud yet, only the title of the
		window just switched to, and cancelling that is what cut it off half way.

		Scoped as tightly as it can be: only while one of the hooked document calls is
		on the stack, and only while a window is being announced. Every other cancel in
		NVDA, including the one the foreground change itself makes, is untouched.
		"""

		def cancelSpeech(*args, **kwargs):
			if getMode() != MODE_NORMAL and _inCallDepth > 0 and _announcingNow():
				return
			return original(*args, **kwargs)

		cancelSpeech.__name__ = "cancelSpeech"
		cancelSpeech.__doc__ = getattr(original, "__doc__", None)
		return cancelSpeech

	def _shouldDropSpellCheckOkObjectSpeech(self, args, kwargs):
		"""Drop NVDA's separate automatic OK-button speech after completion."""
		if time.monotonic() >= self._suppressSpellOkUntil:
			return False
		obj = kwargs.get("obj", args[0] if args else None)
		if obj is None or not _isInOutlookWindow(obj):
			return False
		if "reason" in kwargs:
			reason = kwargs["reason"]
		elif len(args) > 1:
			reason = args[1]
		else:
			return False
		if reason not in _AUTOMATIC_REASONS:
			return False
		try:
			return (
				obj.role == controlTypes.Role.BUTTON
				and (getattr(obj, "name", "") or "").strip().lower() in ("ok", "&ok")
			)
		except Exception:
			return False

	def _makeSpeakObjectWrapper(self, original):
		def speakObject(*args, **kwargs):
			if self._shouldDropSpellCheckOkObjectSpeech(args, kwargs):
				return
			if _shouldDropObjectSpeech(args, kwargs):
				return
			return original(*args, **kwargs)

		speakObject.__name__ = "speakObject"
		speakObject.__doc__ = getattr(original, "__doc__", None)
		return speakObject

	@scriptHandler.script(
		# Translators: Description of a command, shown in the Input Gestures dialog.
		description=_("Cycles the mute browse mode setting between silence, tones and normal"),
	)
	def script_cycleMuteBrowseMode(self, gesture):
		index = (_selectionForMode(getMode()) + 1) % len(MODES)
		setMode(MODES[index])
		_resetGates()
		ui.message(getModeLabels()[index])

	@scriptHandler.script(
		# Translators: Description of a command, shown in the Input Gestures dialog.
		description=_("Cycles the page summary setting between speak, tones and normal"),
	)
	def script_cyclePageSummary(self, gesture):
		index = (_selectionForSummaryMode(getSummaryMode()) + 1) % len(SUMMARY_MODES)
		setSummaryMode(SUMMARY_MODES[index])
		ui.message(getSummaryModeLabels()[index])

	@scriptHandler.script(
		description=_(
			# Translators: Description of a command, shown in the Input Gestures dialog.
			"In a web browser, opens NVDA's find instead of the browser's own find bar. "
			"In Microsoft Outlook it stays Forward. Anywhere else control+F reaches the "
			"program untouched",
		),
	)
	def script_browserFind(self, gesture):
		"""Control+F opens NVDA's find, in a browse mode document or in an edit box.

		Bound to control+F whenever NVDA has something to search, and also whenever a
		browse mode document is present but not wanted for searching — Microsoft
		Outlook's message body, most notably — so that this script, not browse mode's
		own default control+F binding, is the one that runs. See
		L{_syncBrowserFindBinding} for why the binding comes and goes rather than
		staying put, and L{_findSource} for what counts as something to search.

		Where NVDA cannot search, the key is sent back to the program explicitly, so it
		does whatever it always did — Forward, in Outlook.
		"""
		target = _findTarget()
		if target is None:
			# The key belongs to the program, and reaching this at all means the binding
			# has outlived whatever it was made for — Outlook most of all, where control+F
			# is Forward. So put the binding down before handing the key on, and the next
			# press is a real keystroke rather than another injected one.
			log.debug("Mute Browse Mode: nothing to find in, control+F goes to the program")
			self._syncBrowserFindBinding()
			gesture.send()
			return
		log.debug("Mute Browse Mode: opening NVDA's find on %r" % target)
		target.script_find(gesture)

	def _syncBrowserFindBinding(self):
		"""Claim control+F only where NVDA actually has a document to search.

		A bound key is trapped. ``keyboardHandler.internal_keyDownEvent`` returns False
		the moment ``executeGesture`` finds a script for it, so the real key down never
		reaches the program and the matching key up is swallowed too, by way of
		``trappedKeys``. In a browser that is exactly what is wanted, because the whole
		point is to take control+F off the browser. Everywhere else it is exactly what
		is not wanted: control+F is Forward in Microsoft Outlook, and it has to arrive
		there as the keystroke the user actually pressed.

		Hence binding and unbinding as the program in front changes, rather than leaving
		the binding in place and handing the key back with ``gesture.send()``. That would
		put a synthetic keystroke in front of every other program on the system, and a
		synthetic one is not the same as a real one: ``send()`` drops any modifier
		``winUser.getKeyState`` reports as already down, a held control+F auto repeats
		into one injection per repeat, and the 10 ms window ``send()`` allows for NVDA to
		recognise its own injection can be missed under load, after which NVDA reads the
		injection back as a real keystroke.

		The test is the same one the script itself makes: is there a browse mode document
		here that NVDA's find could search, in a program these settings cover. Keying the
		binding to that rather than to the program alone is what keeps the key off
		everything else. Microsoft Word has no such document, so control+F there stays
		Word's own find and arrives as a real keystroke, and so does control+F in the
		browser's address bar or in its find bar — none of them are ever bound, so none
		of them are ever injected.

		Re-checked on every focus change as well as every foreground change, so that the
		binding is in place well before the user can reach for the key.

		Releasing the binding outright is only safe where nothing else is waiting to
		catch control+F once this add-on lets go of it. That holds in Microsoft Word,
		say, which has no browse mode document and so nothing at this layer to catch
		the key at all — control+F reaches Word as the keystroke it always was. It does
		not hold in Microsoft Outlook: the message body is a browse mode document too,
		and browse mode binds its own control+F to NVDA's find by default, the same as
		any web page. Letting go there does not hand the key to Outlook, it exposes
		that default binding underneath — which is NVDA's find opening exactly where
		Forward was wanted. So the binding stays claimed whenever a browse mode
		document is present, whether or not these settings want NVDA's find to open on
		it; L{script_browserFind} sends the real key through itself in that case,
		which is the one situation this add-on cannot avoid ``send()`` for.
		"""
		try:
			focus = api.getFocusObject()
		except Exception:
			focus = None
		try:
			wantsOurFind = _findSource() is not None
		except Exception:
			wantsOurFind = False
		hasDocumentToGuard = focus is not None and _activeBrowseModeTreeInterceptor(focus) is not None
		wanted = wantsOurFind or hasDocumentToGuard
		if wanted == self._browserFindBound:
			return
		try:
			if wanted:
				self.bindGesture("kb:control+f", "browserFind")
			else:
				self.removeGestureBinding("kb:control+f")
			self._browserFindBound = wanted
			log.debug("Mute Browse Mode: control+F %s" % ("claimed" if wanted else "released"))
			if _tracing:
				# Only when the answer changes, and only on the main thread. Which program
				# the add-on thinks it is in is the whole of this decision, and the one
				# thing a keystroke trace cannot show on its own.
				_traceWrite(
					"FIND control+F %s | outlookIsCurrent=%s"
					% ("claimed" if wanted else "released", outlookIsCurrent()),
				)
		except Exception:
			log.error("Mute Browse Mode: could not change the control+F binding", exc_info=True)
