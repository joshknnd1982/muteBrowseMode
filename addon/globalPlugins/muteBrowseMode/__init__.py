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

import time

import addonHandler
import api
import browseMode
import config
import controlTypes
import core
import globalPluginHandler
import inputCore
import scriptHandler
import speech
import speech.speech
import tones
import ui
import virtualBuffers
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


def _targetOf(obj):
	"""Which of the two applications C{obj} belongs to, or C{None} for anything else."""
	if obj is None:
		return None
	appName = _appNameOf(obj)
	if appName in _OUTLOOK_APP_NAMES:
		return TARGET_OUTLOOK
	if appName in _CHROMIUM_APP_NAMES:
		return TARGET_CHROMIUM
	if _windowClassOf(obj) in _CHROMIUM_WINDOW_CLASSES:
		return TARGET_CHROMIUM
	return None


def _isOutlook(obj):
	"""Whether C{obj} belongs to any version of Microsoft Outlook."""
	return obj is not None and _appNameOf(obj) in _OUTLOOK_APP_NAMES


def outlookIsCurrent(obj=None):
	"""Whether Microsoft Outlook is the program the user is currently in.

	Answered from the foreground window and the focus, so it is the running program
	that decides, not the object that happened to raise an event. C{obj}, when given,
	is checked first: a document belonging to Outlook counts as Outlook even in the
	moment before its window has become the foreground one.

	The answer is only ever written to NVDA's log. Nothing here reaches the
	synthesiser, so the check itself is never heard.
	"""
	if _isOutlook(obj):
		return True
	for getter in (api.getForegroundObject, api.getFocusObject):
		try:
			candidate = getter()
		except Exception:
			continue
		if _isOutlook(candidate):
			return True
	return False


def _isWebBrowser(obj):
	"""Whether C{obj} belongs to a browser or an Electron application, and not Outlook.

	Outlook is excluded before anything else, because the new Outlook for Windows is a
	WebView2 application and so looks exactly like Chromium from the outside.
	"""
	if obj is None or _isOutlook(obj):
		return False
	if _targetOf(obj) == TARGET_CHROMIUM:
		return True
	if _appNameOf(obj) in _GECKO_APP_NAMES:
		return True
	return _windowClassOf(obj) in _GECKO_WINDOW_CLASSES


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

#: Roles used for toasts, flash messages and other transient announcements.
_ALERT_ROLES = _members(controlTypes.Role, ("ALERT", "TOOLTIP", "HELPBALLOON", "NOTIFICATION"))

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

#: Window classes Outlook renders a message body into: the Word editor it composes
#: rich text in, the HTML control, and the plain text one.
_OUTLOOK_BODY_WINDOW_CLASSES = frozenset((
	"_WwG",
	"_WwB",
	"Internet Explorer_Server",
	"RichEdit20W",
	"RICHEDIT50W",
))

#: Roles a message body can have. Everything else the user tabs through on the way to
#: it, To, Cc, Subject and the toolbars, has some other role or is a single line.
_BODY_ROLES = _members(controlTypes.Role, ("DOCUMENT", "EDITABLETEXT"))

_STATE_READONLY = getattr(controlTypes.State, "READONLY", None)
_STATE_UNAVAILABLE = getattr(controlTypes.State, "UNAVAILABLE", None)
_STATE_MULTILINE = getattr(controlTypes.State, "MULTILINE", None)
_ROLE_DOCUMENT = getattr(controlTypes.Role, "DOCUMENT", None)


def _isOutlookMessageBody(obj):
	"""Whether the focus has landed in an Outlook message body that can be typed into.

	The address and subject fields are editable text too, so a plain "is this editable"
	test is not enough to tell them apart. What separates the body from them is that it
	takes more than one line, or that it is a whole document rather than a field, or
	that it is one of the windows Outlook only ever puts a message body in. A body the
	user cannot type into, such as the one in the reading pane, is not announced,
	because the announcement invites them to type.
	"""
	if not _isOutlook(obj):
		return False
	role = getattr(obj, "role", None)
	if role not in _BODY_ROLES:
		return False
	try:
		states = set(obj.states or ())
	except Exception:
		states = set()
	if _STATE_READONLY is not None and _STATE_READONLY in states:
		return False
	if _STATE_UNAVAILABLE is not None and _STATE_UNAVAILABLE in states:
		return False
	if _windowClassOf(obj) in _OUTLOOK_BODY_WINDOW_CLASSES:
		return True
	if role == _ROLE_DOCUMENT:
		return True
	return _STATE_MULTILINE is not None and _STATE_MULTILINE in states


def _announceMessageBody():
	# Translators: Announced when the focus reaches the message body in Microsoft Outlook.
	with _ownSpeech():
		ui.message(_("You are now in the message body, type a message."))


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


def _shouldDropObjectSpeech(args, kwargs):
	"""True for the window, dialog and document titles and toasts we never want spoken.

	Only applies inside Outlook and Chromium based browsers, and only to announcements
	NVDA made on its own initiative.
	"""
	if getMode() == MODE_NORMAL:
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
	return _targetOf(obj) is not None


def _onGesture(*args, **kwargs):
	"""Any key the user presses means they want to hear things again.

	Registered with the decide_executeGesture extension point purely to get told
	about input; it never vetoes a gesture.
	"""
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


def _makeSettingsWrapper(original):
	def makeSettings(self, settingsSizer):
		original(self, settingsSizer)
		try:
			_addChoices(self, guiHelper.BoxSizerHelper(self, sizer=settingsSizer))
		except Exception:
			self._muteBrowseModeChoice = None
			self._muteBrowseModeSummaryChoice = None
			self._muteBrowseModeLoadingCheckBox = None
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
		except Exception:
			log.error("Mute Browse Mode: could not install speech hooks", exc_info=True)
			_unpatchAll()
			return

		# Each of the remaining hooks is optional and independent: if one of them does
		# not fit this NVDA, the rest still do their job.
		self._installDocumentHooks()
		self._installNoiseHooks()

		try:
			inputCore.decide_executeGesture.register(_onGesture)
			self._gestureHandlerRegistered = True
		except Exception:
			log.error("Mute Browse Mode: could not hook input gestures", exc_info=True)

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
		try:
			isOutlook = _isOutlook(obj)
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
		if time.monotonic() < self._outlookArrivalUntil and _isOutlook(obj):
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

	def _makeSpeakWrapper(self, original):
		def speak(*args, **kwargs):
			if getMode() != MODE_NORMAL and _isGated():
				return
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

	def _makeSpeakObjectWrapper(self, original):
		def speakObject(*args, **kwargs):
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
			"Passes control+F straight to the program, so that it forwards the current "
			"message in Microsoft Outlook. NVDA's own find stays on NVDA+control+F",
		),
		gesture="kb:control+f",
	)
	def script_passThroughControlF(self, gesture):
		"""Make sure control+F belongs to the application and never to NVDA.

		In Microsoft Outlook control+F is Forward, and it has to reach Outlook for the
		message to be forwarded. Binding it here rather than leaving it unbound is what
		guarantees that: a global plugin script is the first thing NVDA looks at, ahead
		of the browse mode document and everything else that might otherwise claim the
		key, and this one does nothing but hand the keystroke on. NVDA ignores the keys
		it injects itself, so this cannot come back round.

		Nothing is lost elsewhere: control+F reaching the program is exactly what a
		browser's or an editor's find bar wants too. NVDA's find is unaffected, because
		NVDA binds that to NVDA+control+F.
		"""
		if outlookIsCurrent():
			log.debug("Mute Browse Mode: passing control+F to Outlook to forward the message")
		gesture.send()
