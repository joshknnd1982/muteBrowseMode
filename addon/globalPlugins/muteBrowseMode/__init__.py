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

#: Monotonic timestamp at which speech is allowed through again. 0 means open.
_gateUntil = 0.0
#: Depth of nested "silence just this call" wrappers, and the latest moment any of them
#: may keep speech shut. Both have to agree, so even a lost decrement expires by itself.
_muteDepth = 0
_muteUntil = 0.0
#: Depth of nested "this is the add-on speaking" wrappers. Speech from inside one of
#: these is never gated: it is the add-on's own announcement, not NVDA's.
_bypassDepth = 0


def _openGate(seconds):
	global _gateUntil
	_gateUntil = time.monotonic() + seconds


def _closeGate():
	global _gateUntil
	_gateUntil = 0.0


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
	global _muteDepth, _muteUntil, _bypassDepth
	_muteDepth = 0
	_muteUntil = 0.0
	_bypassDepth = 0
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
	if _gateUntil <= 0.0 or now >= _gateUntil:
		return False
	if _sayAllRunning():
		# Say all is an explicit "read this to me" request, including read on page
		# load. Never swallow it, and close the gate so its start is not clipped.
		_closeGate()
		return False
	return True


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


def _hookGate(owner, name, inCall, after, chime=False, onlyIf=None, chimeCheck=None, post=None):
	"""Hold the speech gate open across C{owner.name}.

	@param inCall: seconds the gate is held while the call is on the stack.
	@param after: seconds the gate is held once the call returns.
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
			_openGate(inCall)
		try:
			return original(self, *args, **kwargs)
		finally:
			if silence:
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


def _refreshChoices(panel):
	choice = getattr(panel, "_muteBrowseModeChoice", None)
	if choice is not None:
		choice.SetSelection(_selectionForMode(getMode()))
	choice = getattr(panel, "_muteBrowseModeSummaryChoice", None)
	if choice is not None:
		choice.SetSelection(_selectionForSummaryMode(getSummaryMode()))


def _makeSettingsWrapper(original):
	def makeSettings(self, settingsSizer):
		original(self, settingsSizer)
		try:
			_addChoices(self, guiHelper.BoxSizerHelper(self, sizer=settingsSizer))
		except Exception:
			self._muteBrowseModeChoice = None
			self._muteBrowseModeSummaryChoice = None
			log.error(
				"Mute Browse Mode: could not add the combo boxes to Browse Mode settings",
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
					inCall=_IN_CALL_GATE,
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
				dict(inCall=_IN_CALL_GATE, after=_TRAILING_GATE, chime=True),
			),
			# A virtual buffer starting to load, which is what silences "Loading
			# document...", and the point the page summary is owed from.
			(
				virtualBuffers.VirtualBuffer,
				"loadBuffer",
				dict(inCall=_LOAD_GATE, after=_LOAD_GATE, post=_armSummary),
			),
			# ...and finishing, which speaks "Refreshed" on a reload and is the moment
			# both the chime and the page summary belong to.
			(
				virtualBuffers.VirtualBuffer,
				"_loadBufferDone",
				dict(
					inCall=_IN_CALL_GATE,
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
