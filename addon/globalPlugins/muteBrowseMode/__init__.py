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
window or document titles, toasts, or live region "flash" messages, because in those
two applications they are noise rather than information.

The gate is deadline based rather than a counter, so a bug or an exception can never
leave NVDA permanently mute: the worst case is a few seconds of silence that expires
on its own. Any input gesture closes the gate immediately, so pressing a key always
gets normal speech back straight away. Braille is deliberately untouched: everything
here suppresses speech only, and the calls that would have spoken still run, so
braille, vision and NVDA's own caches carry on as usual.
"""

import time

import addonHandler
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


#: Section this add-on owns in nvda.ini.
CONF_SECTION = "muteBrowseMode"

MODE_SILENCE = "silence"
MODE_TONES = "tones"
MODE_NORMAL = "normal"

#: Order matters: this is the order of the entries in the combo box.
MODES = (MODE_SILENCE, MODE_TONES, MODE_NORMAL)

config.conf.spec[CONF_SECTION] = {
	"mode": 'option("silence", "tones", "normal", default="normal")',
}


def getModeLabels():
	"""Combo box entries, in the same order as L{MODES}."""
	return [
		# Translators: A choice in the "Mute browse mode" combo box in Speech settings.
		_("Silence all browsing"),
		# Translators: A choice in the "Mute browse mode" combo box in Speech settings.
		_("Play tones"),
		# Translators: A choice in the "Mute browse mode" combo box in Speech settings.
		_("Normal"),
	]


def getMode():
	"""The configured mode, falling back to "normal" if the config is missing or odd."""
	try:
		mode = config.conf[CONF_SECTION]["mode"]
	except Exception:
		return MODE_NORMAL
	return mode if mode in MODES else MODE_NORMAL


def setMode(mode):
	config.conf[CONF_SECTION]["mode"] = mode


### The applications we are aggressive in

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
_CHROMIUM_WINDOW_CLASSES = frozenset((
	"Chrome_WidgetWin_0",
	"Chrome_WidgetWin_1",
	"Chrome_RenderWidgetHostHWND",
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


### Roles and reasons

def _members(enum, names):
	"""The named members of C{enum} that this NVDA actually has."""
	return frozenset(m for m in (getattr(enum, name, None) for name in names) if m is not None)


#: Roles whose name is really just a window or document title. NVDA speaks these when a
#: window comes to the foreground and as it walks down the focus ancestors, which is
#: where the Outlook message list window title and the message window title come from.
_TITLE_ROLES = _members(
	controlTypes.Role,
	("WINDOW", "PANE", "FRAME", "INTERNALFRAME", "DOCUMENT", "APPLICATION", "PROPERTYPAGE"),
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


def _resetGates():
	global _muteDepth, _muteUntil
	_muteDepth = 0
	_muteUntil = 0.0
	_closeGate()


def _sayAllRunning():
	try:
		from speech.sayAll import SayAllHandler

		return SayAllHandler.isRunning()
	except Exception:
		return False


def _isGated():
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


### Monkey patching

#: (owner, name, original, wasOwnAttribute, replacement) for everything we patched.
_patches = []


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


def _hookGate(owner, name, inCall, after, chime=False, onlyIf=None, chimeCheck=None):
	"""Hold the speech gate open across C{owner.name}.

	@param inCall: seconds the gate is held while the call is on the stack.
	@param after: seconds the gate is held once the call returns.
	@param chime: play the ready tones on the way out, in "play tones" mode.
	@param onlyIf: optional callable(self, args, kwargs) deciding whether this
		particular call is one of the ones we silence. Checked before the original
		runs, because some of what it looks at is cleared by the original itself.
	@param chimeCheck: optional callable(self, args, kwargs) vetoing only the chime.
	"""
	original = getattr(owner, name)

	def wrapper(self, *args, **kwargs):
		mode = getMode()
		if mode == MODE_NORMAL or (onlyIf is not None and not onlyIf(self, args, kwargs)):
			return original(self, *args, **kwargs)
		_openGate(inCall)
		try:
			return original(self, *args, **kwargs)
		finally:
			_openGate(after)
			if chime and mode == MODE_TONES and (chimeCheck is None or chimeCheck(self, args, kwargs)):
				_playReadyTones()

	wrapper.__name__ = name
	wrapper.__doc__ = getattr(original, "__doc__", None)
	# eventHandler reads flags such as ignoreIsReady off the handler it is about to
	# call, so anything the original carried has to survive the wrapping.
	for flag in ("ignoreIsReady",):
		if hasattr(original, flag):
			setattr(wrapper, flag, getattr(original, flag))
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

	wrapper.__name__ = name
	wrapper.__doc__ = getattr(original, "__doc__", None)
	_patch(owner, name, wrapper)


def _loadSucceeded(self, args, kwargs):
	"""_loadBufferDone(self, success=True): don't chime for a failed load."""
	if "success" in kwargs:
		return bool(kwargs["success"])
	if args:
		return bool(args[0])
	return True


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
	"""True for the window titles, document titles and toasts we never want spoken.

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
	return True


### Speech settings panel

def _selectionForMode(mode):
	try:
		return MODES.index(mode)
	except ValueError:
		return MODES.index(MODE_NORMAL)


def _makeSettingsWrapper(original):
	def makeSettings(self, settingsSizer):
		original(self, settingsSizer)
		try:
			sHelper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
			self._muteBrowseModeChoice = sHelper.addLabeledControl(
				# Translators: Label of a combo box added to NVDA's Speech settings.
				_("&Mute browse mode:"),
				wx.Choice,
				choices=getModeLabels(),
			)
			self._muteBrowseModeChoice.SetSelection(_selectionForMode(getMode()))
		except Exception:
			self._muteBrowseModeChoice = None
			log.error("Mute Browse Mode: could not add the combo box to Speech settings", exc_info=True)

	return makeSettings


def _onSaveWrapper(original):
	def onSave(self):
		original(self)
		choice = getattr(self, "_muteBrowseModeChoice", None)
		if choice is None:
			return
		index = choice.GetSelection()
		if 0 <= index < len(MODES):
			setMode(MODES[index])

	return onSave


def _onPanelActivatedWrapper(original):
	def onPanelActivated(self):
		# NVDA reuses the settings dialog, so refresh in case the mode was changed
		# elsewhere (a profile switch, or the cycle script) since it was built.
		choice = getattr(self, "_muteBrowseModeChoice", None)
		if choice is not None:
			choice.SetSelection(_selectionForMode(getMode()))
		original(self)

	return onPanelActivated


def _patchSpeechPanel():
	panel = settingsDialogs.SpeechSettingsPanel
	_patch(panel, "makeSettings", _makeSettingsWrapper(panel.makeSettings))
	_patch(panel, "onSave", _onSaveWrapper(panel.onSave))
	_patch(panel, "onPanelActivated", _onPanelActivatedWrapper(panel.onPanelActivated))


### Fallback settings panel

class MuteBrowseModePanel(settingsDialogs.SettingsPanel):
	"""Only registered if the combo box could not be injected into the Speech panel."""

	# Translators: Title of the add-on's own settings category, used as a fallback.
	title = _("Mute Browse Mode")

	def makeSettings(self, settingsSizer):
		sHelper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		self._muteBrowseModeChoice = sHelper.addLabeledControl(
			# Translators: Label of a combo box added to NVDA's Speech settings.
			_("&Mute browse mode:"),
			wx.Choice,
			choices=getModeLabels(),
		)
		self._muteBrowseModeChoice.SetSelection(_selectionForMode(getMode()))

	def onSave(self):
		index = self._muteBrowseModeChoice.GetSelection()
		if 0 <= index < len(MODES):
			setMode(MODES[index])


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	# Translators: Category for this add-on's commands in the Input Gestures dialog.
	scriptCategory = _("Mute Browse Mode")

	def __init__(self):
		super().__init__()
		self._speechPanelPatched = False
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
			_patchSpeechPanel()
			self._speechPanelPatched = True
		except Exception:
			log.error(
				"Mute Browse Mode: could not extend the Speech settings panel, "
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
			# document...".
			(
				virtualBuffers.VirtualBuffer,
				"loadBuffer",
				dict(inCall=_LOAD_GATE, after=_LOAD_GATE),
			),
			# ...and finishing, which speaks "Refreshed" on a reload and is the moment
			# the chime belongs to.
			(
				virtualBuffers.VirtualBuffer,
				"_loadBufferDone",
				dict(
					inCall=_IN_CALL_GATE,
					after=_TRAILING_GATE,
					chime=True,
					chimeCheck=_loadSucceeded,
				),
			),
		)
		for owner, name, options in hooks:
			try:
				_hookGate(owner, name, **options)
			except Exception:
				log.error("Mute Browse Mode: could not hook %s.%s" % (owner.__name__, name), exc_info=True)

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
