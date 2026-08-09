# -*- coding: utf-8 -*-
# Mute Browse Mode, an NVDA add-on.
# Copyright (C) 2026 Josh Kennedy
# This file is covered by the GNU General Public License, version 2.

"""Stops NVDA announcing a browse mode document every time one loads or is entered.

NVDA normally speaks the document name, the word "document" and the first line of the
buffer whenever a web page finishes loading or an Outlook message is opened. This
add-on gates speech across those moments instead, and can play a short ascending
chime in its place so there is still a cue that the document is ready.

The gate is deadline based rather than a counter, so a bug or an exception can never
leave NVDA permanently mute: the worst case is a few seconds of silence that expires
on its own. Any input gesture closes the gate immediately, so pressing a key always
gets normal speech back straight away.
"""

import time

import addonHandler
import browseMode
import config
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


### The speech gate

#: Ceiling while a hooked call is still on the stack.
_IN_CALL_GATE = 5.0
#: How long the gate is held after a hooked call returns. NVDA queues part of the
#: document announcement onto the main queue, so it does not all happen inline.
_TRAILING_GATE = 0.5
#: Ceiling while a virtual buffer is loading. This is what silences "Loading
#: document...", which NVDA speaks from a timer part way through the load.
_LOAD_GATE = 10.0

#: Monotonic timestamp at which speech is allowed through again. 0 means open.
_gateUntil = 0.0


def _openGate(seconds):
	global _gateUntil
	_gateUntil = time.monotonic() + seconds


def _closeGate():
	global _gateUntil
	_gateUntil = 0.0


def _sayAllRunning():
	try:
		from speech.sayAll import SayAllHandler

		return SayAllHandler.isRunning()
	except Exception:
		return False


def _isGated():
	if _gateUntil <= 0.0 or time.monotonic() >= _gateUntil:
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
#: Loading a buffer fires more than one of our hooks. Only chime once per document.
_TONE_DEBOUNCE = 0.75

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


def _hookGate(owner, name, inCall, after, chime=False, chimeCheck=None):
	"""Hold the speech gate open across C{owner.name}.

	@param inCall: seconds the gate is held while the call is on the stack.
	@param after: seconds the gate is held once the call returns.
	@param chime: play the ready tones on the way out, in "play tones" mode.
	@param chimeCheck: optional callable(args, kwargs) vetoing the chime.
	"""
	original = getattr(owner, name)

	def wrapper(self, *args, **kwargs):
		mode = getMode()
		if mode == MODE_NORMAL:
			return original(self, *args, **kwargs)
		_openGate(inCall)
		try:
			return original(self, *args, **kwargs)
		finally:
			_openGate(after)
			if chime and mode == MODE_TONES and (chimeCheck is None or chimeCheck(args, kwargs)):
				_playReadyTones()

	wrapper.__name__ = name
	wrapper.__doc__ = getattr(original, "__doc__", None)
	_patch(owner, name, wrapper)


def _loadSucceeded(args, kwargs):
	"""_loadBufferDone(self, success=True): don't chime for a failed load."""
	if "success" in kwargs:
		return bool(kwargs["success"])
	if args:
		return bool(args[0])
	return True


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
			# Gating speech itself, rather than each announcement, catches every
			# route into the synthesiser: speakTextInfo, speakObject, speakMessage
			# and ui.message all funnel through speech.speech.speak.
			speakWrapper = self._makeSpeakWrapper(speech.speech.speak)
			_patch(speech.speech, "speak", speakWrapper)
			_patch(speech, "speak", speakWrapper)

			# Entering any browse mode document: web pages, and Outlook / Word
			# message bodies, which are browse mode documents but not virtual buffers.
			_hookGate(
				browseMode.BrowseModeDocumentTreeInterceptor,
				"event_treeInterceptor_gainFocus",
				inCall=_IN_CALL_GATE,
				after=_TRAILING_GATE,
				chime=True,
			)
			# A virtual buffer starting to load, which is what silences
			# "Loading document...".
			_hookGate(
				virtualBuffers.VirtualBuffer,
				"loadBuffer",
				inCall=_LOAD_GATE,
				after=_LOAD_GATE,
			)
			# ...and finishing, which is the moment the chime belongs to.
			_hookGate(
				virtualBuffers.VirtualBuffer,
				"_loadBufferDone",
				inCall=_IN_CALL_GATE,
				after=_TRAILING_GATE,
				chime=True,
				chimeCheck=_loadSucceeded,
			)
		except Exception:
			log.error("Mute Browse Mode: could not install speech hooks", exc_info=True)
			_unpatchAll()
			return

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

	def terminate(self):
		_closeGate()
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
		return speak

	@scriptHandler.script(
		# Translators: Description of a command, shown in the Input Gestures dialog.
		description=_("Cycles the mute browse mode setting between silence, tones and normal"),
	)
	def script_cycleMuteBrowseMode(self, gesture):
		index = (_selectionForMode(getMode()) + 1) % len(MODES)
		setMode(MODES[index])
		_closeGate()
		ui.message(getModeLabels()[index])
