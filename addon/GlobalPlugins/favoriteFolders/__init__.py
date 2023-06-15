#-*- coding: utf-8 -*-
# Favorite Folders add-on for NVDA.
# Registers the most used folders.
# It can helps in the process of saving files, because you can enter the address at the beginning of a editable text line.
# Shortcut: WINDOWS+BACKSPACE
# written by Rui Fontes, and Ângelo Abrantes based on a work of Marcos Antonio de Oliveira.
# Copyright (C) 2020-2023 Rui Fontes <rui.fontes@tiflotecnia.com>
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.

# Import the necessary modules
import globalPluginHandler
import globalVars
import gui
import wx
import os
import api
import ui
import re
import windowUtils
import controlTypes
from ctypes import windll
from sys import getwindowsversion
import NVDAObjects
from keyboardHandler import KeyboardInputGesture
import winUser
from . import win32con
from configobj import ConfigObj
from time import sleep
from scriptHandler import script
from threading import Thread, Event
# For translation process
import addonHandler
# To start the translation process
addonHandler.initTranslation()

# globalVars
#_ffIniFile = os.path.join (os.path.dirname(__file__),'FavoriteFolders.ini')
_ffIniFile = os.path.abspath(os.path.join(globalVars.appArgs.configPath, "FavoriteFolders.ini"))
# Translators: Title of add-on, present in the dialog boxes.
title = _("Favorite folders")
showPath = False
dictFolders = {}
newFolder = ""


class TimeoutThread(Thread):
	def __init__(self, target, args):
		super().__init__(target=target, args=args, daemon=True)
		self.target = target
		self.args = args

	def run_with_timeout(self, timeout):
		finish_event = Event()

		def helper():
			self.run()
			finish_event.set()

		helper_thread = Thread(target=helper)
		helper_thread.start()
		helper_thread.join(timeout)

		if helper_thread.is_alive():
			pass
		else:
			return self.args[0] #self.target(*self.args)


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	def __init__(self):
		super(globalPluginHandler.GlobalPlugin, self).__init__()
		self.lastForeground = 0
		self.dialog = None

	def check_path(self, path):
		try:
			driveLetter = path [:3]
			driveType = windll.kernel32.GetDriveTypeW(driveLetter)
			if driveType == 3: # fixed disk
				self.result = path
			elif not os.path.isdir(path):
				self.result = None
			else:
				self.result = path
		except:
			self.result = path

	def readConfig (self):
		if not os.path.isfile (_ffIniFile):
			return None
		config = ConfigObj(_ffIniFile, encoding = "utf-8")
		if "options" in config.sections:
			try:
				showPath = config['Options']['ShowPath'] == '1'
			except:
				showPath = False
		try:
			folders = config['Folders']
		except KeyError:
			folders = {}
		total = len(folders.keys())
		if not total:
			return None
		for item in folders.keys ():
			try:
				path = folders[item].decode('utf-8')
			except:
				path = folders[item]
			thread = TimeoutThread(target=self.check_path, args=(path,))
			result = thread.run_with_timeout(timeout = 0.001)
			if result is None or not result:
				folders.__delitem__(item)
			elif self.result is None:
				folders.__delitem__(item)
		if not len (folders.keys()):
			folders = None
		global dictFolders
		dictFolders = folders
		return folders

	@script(
	# For translators: Message to be announced during Keyboard Help
	description = _("Opens a dialog box to register and open favorite folders."),
	gesture = "kb:WINDOWS+Backspace")
	def script_startFavoriteFolders(self, gesture):
		path = ' '
		focusObj = api.getFocusObject()
		# If the focus is on windows explorer, gets the address of the folder.
		if 'explorer' in focusObj.appModule.appModuleName:
			hForeground = api.getForegroundObject().windowHandle
			h = self.findDescendantWindow(hForeground, 1001) 
			if not h:
				h = self.findDescendantWindow(hForeground, 41477)
				h = self.findDescendantWindow(h, "Edit")
			if h:
				obj = NVDAObjects.IAccessible.getNVDAObjectFromEvent (h, -4, 0)
				if getwindowsversion().major == 5: # windows xp.
					path = obj.value
				else:
					name = obj.name
					pattern = re.compile ('\w:\\\\.{0,}')
					result = re.findall (pattern, name)
					if result !=[]:
						path = result [0]
				if path[-1] != '\\':
					path += '\\'
				if  not os.path.isdir (path):
					path = None
		else:
			path = None
		global newFolder
		newFolder = path
		self.showFavoriteFoldersDialog ()

	def findDescendantWindow (self, h, c):
		# Returns the handle of a window with a given class or control id sought from a specified window.
		try:
			if type (c) is int:
				return windowUtils.findDescendantWindow (h, controlID=c)
			else:
				return windowUtils.findDescendantWindow (h, className=c)
		except LookupError:
			return False

	def showFavoriteFoldersDialog (self):
		global newFolder
		# Displays the add-on dialog box.
		dictFolders = self.readConfig()
		if dictFolders is None and newFolder is None:
			# Translators: Announced when the focus is not on the Windows Explorer and there is no registered folders.
			ui.message (_('You do not have added folders and is not in the windows explorer window to make a record'))
			return
		elif dictFolders is not None:
			values = dictFolders.values()
			# if the focus is on Windows Explorer, check if the folder is already registered.
			if newFolder is not None:
				if   (newFolder.lower() in [obj.lower() for obj in values]):
					newFolder = None
		self.lastForeground = api.getForegroundObject().windowHandle
		global lastFocus, lastValue
		lastFocus = api.getFocusObject()
		lastValue = api.getFocusObject().value
		self.dialog = FavoriteFoldersDialog (gui.mainFrame)
		self.dialog.updateFolders(dictFolders,0)
		self.dialog.setButtons()
		if not self.dialog.IsShown():
			gui.mainFrame.prePopup()
			self.dialog.Show()
			self.dialog.CentreOnScreen()
			gui.mainFrame.postPopup()

	def terminate (self):
		if self.dialog is not None:
			self.dialog.Destroy()

	def event_gainFocus (self, obj, nextHandler):
		# Puts focus on the edit box, because in some applications, such as Skype, focus is lost in open file dialog.
		if self.dialog is not None:
			hRealWindow = winUser.getAncestor (obj.windowHandle, 2)
			if self.dialog.writtenAddress == True and hRealWindow == self.lastForeground and lastFocus != obj and obj.windowHandle != hRealWindow:
				winUser.sendMessage(api.getForegroundObject().windowHandle, win32con.WM_NEXTDLGCTL ,lastFocus.windowHandle,1)
				api.setMouseObject(lastFocus)
				api.moveMouseToNVDAObject(lastFocus)
			self.dialog.writtenAddress = False	

		nextHandler()


# To avoid use on secure screens
if globalVars.appArgs.secure:
	# Override the global plugin to disable it.
	GlobalPlugin = globalPluginHandler.GlobalPlugin


class FavoriteFoldersDialog(wx.Dialog):
	def __init__(self, *args, **kwds):
		kwds["style"] = kwds.get("style", 0) | wx.DEFAULT_DIALOG_STYLE
		wx.Dialog.__init__(self, *args, **kwds)
		self.dialogActive = False
		self.writtenAddress = False
		self.SetTitle(title)


		sizer_1 = wx.BoxSizer(wx.VERTICAL)

		label_1 = wx.StaticText(self, wx.ID_ANY, _("Favorite folders list"))
		sizer_1.Add(label_1, 0, 0, 0)

		self.listBox = wx.ListCtrl(self, wx.ID_ANY, style=wx.BORDER_SUNKEN | wx.LC_HRULES | wx.LC_REPORT | wx.LC_SORT_ASCENDING | wx.LC_VRULES)
		sizer_1.Add(self.listBox, 8, wx.EXPAND, 0)

		config = ConfigObj(_ffIniFile, encoding = "utf-8")
		self.chkAddress = wx.CheckBox(self, wx.ID_ANY, _("&Show paths in the list"))
		if "Options" in config.sections:
			print("Certo")
			if config['Options']['ShowPath'] == "1":
				print("1")
				self.chkAddress.SetValue(True)
		sizer_1.Add(self.chkAddress, 0, 0, 0)

		self.addButton = wx.Button(self, wx.ID_ANY, _("&Add folder"))
		sizer_1.Add(self.addButton, 0, 0, 0)

		self.openButton = wx.Button(self, wx.ID_ANY, _("&Open folder"))
		sizer_1.Add(self.openButton, 0, 0, 0)

		self.pastButton = wx.Button(self, wx.ID_ANY, _("Write in the e&dit box"))
		sizer_1.Add(self.pastButton, 0, 0, 0)

		self.renameButton = wx.Button(self, wx.ID_ANY, _("Re&name"))
		sizer_1.Add(self.renameButton, 0, 0, 0)

		self.removeButton = wx.Button(self, wx.ID_ANY, _("&Remove"))
		sizer_1.Add(self.removeButton, 0, 0, 0)

		sizer_2 = wx.StdDialogButtonSizer()
		sizer_1.Add(sizer_2, 0, wx.ALIGN_RIGHT | wx.ALL, 4)

		self.button_CLOSE = wx.Button(self, wx.ID_CLOSE, "")
		sizer_2.AddButton(self.button_CLOSE)

		sizer_2.Realize()

		self.SetSizer(sizer_1)
		sizer_1.Fit(self)

		self.SetEscapeId(self.button_CLOSE.GetId())

		self.Layout()

		self.Bind(wx.EVT_CHECKBOX, self.onCheckAddress, self.chkAddress)
		self.Bind(wx.EVT_BUTTON, self.onAdd, self.addButton)
		self.Bind(wx.EVT_BUTTON, self.onOpen, self.openButton)
		self.Bind(wx.EVT_BUTTON, self.onPast, self.pastButton)
		self.Bind(wx.EVT_BUTTON, self.onRename, self.renameButton)
		self.Bind(wx.EVT_BUTTON, self.onRemove, self.removeButton)
		self.listBox.Bind(wx.EVT_KEY_DOWN, self.onKeyPress)
		wx.EVT_KILL_FOCUS (self.listBox,self.onFocusLost)
		wx.EVT_KILL_FOCUS (self.chkAddress,self.onFocusLost)
		wx.EVT_KILL_FOCUS (self.addButton,self.onFocusLost)
		wx.EVT_KILL_FOCUS (self.openButton,self.onFocusLost)
		wx.EVT_KILL_FOCUS (self.pastButton,self.onFocusLost)
		wx.EVT_KILL_FOCUS (self.renameButton,self.onFocusLost)
		wx.EVT_KILL_FOCUS (self.removeButton,self.onFocusLost)
		wx.EVT_KILL_FOCUS (self.button_CLOSE,self.onFocusLost)

	def onCheckAddress (self, evt):
		# enables or disables the display of addresses of the folders in the list.
		evt.Skip()
		# Saves the status of the check box on the ini file
		config = ConfigObj(_ffIniFile, encoding = "utf-8")
		try:
			config['Options']['ShowPath'] = int (self.chkAddress.GetValue())
			config.write()
		except:
			config['Options'] = {}
			config['Options'] = {'ShowPath':int(self.chkAddress.GetValue())}
			config.write()
		index=self.listBox.GetFocusedItem()
		self.updateFolders(dictFolders, index)

	def onAdd (self, evt):
		# Add a new folder.
		evt.Skip()
		self.dialogActive = True
		# Translators: Message dialog box to add a new folder.
		dlg = wx.TextEntryDialog(gui.mainFrame,_("Enter a nickname for the folder"), title, newFolder)
		def callback (result):
			global dictFolders, newFolder
			if  result == wx.ID_OK:
				nickName = dlg.GetValue().strip().upper()
				if nickName != "":
					if self.listBox.FindItem (0, nickName) == -1:
						config = ConfigObj(_ffIniFile, encoding = "utf-8")
						if 'Folders' in config.sections:
							folders = config['Folders']
							folders.__setitem__ (nickName, newFolder)
						else:
							config['Folders'] = {nickName:newFolder}
							dictFolders = {nickName:newFolder}
						config.write()
						self.listBox.Append([nickName])
						newIndex = self.listBox.FindItem(0,nickName)
						if self.chkAddress.GetValue():
							self.listBox.SetStringItem (newIndex, 1, newFolder)
						List = list(dictFolders.items())
						List.append ((nickName, newFolder))
						dictFolders = dict(List)
						# Puts the focus on the inserted folder.
						self.listBox.Focus (newIndex)
						self.listBox.Select(newIndex)
						self.listBox.SetFocus()
						# Defines the status of the buttons and check box.
						newFolder = ""
						self.setButtons()
					else:
						# Translators: Announcement that the folder name already exists in the list.
						gui.messageBox (_("There is already a folder with this nickname!"), title)
				self.dialogActive = False
		gui.runScriptModalDialog(dlg, callback)

	def onOpen (self, evt):
		# Opens the selected folder.
		self.Hide()
		evt.Skip()
		index=self.listBox.GetFocusedItem()
		nickName = self.listBox.GetItemText(index)
		path = dictFolders [nickName]
		os.startfile(path)

	def onPast (self, evt):
		# Simulates typing the folder address in the edit box.
		self.Hide()
		evt.Skip()
		index=self.listBox.GetFocusedItem()
		nickName = self.listBox.GetItemText(index)
		path = dictFolders [nickName] 
		lastFocus.setFocus()
		# insert path in editable text.
		winUser.sendMessage (lastFocus.windowHandle, win32con.WM_SETTEXT, 0, path + lastValue)
		# Positions the insertion cursor at the end of the line.
		winUser.sendMessage (lastFocus.windowHandle, win32con.WM_KEYDOWN, win32con.VK_END, 1)
		winUser.sendMessage (lastFocus.windowHandle, win32con.WM_KEYUP, win32con.VK_END, 1)
		self.writtenAddress = True

	def onRename (self, evt):
		# Renames the reference of the selected folder.
		evt.Skip()
		global dictFolders, folders
		index=self.listBox.GetFocusedItem()
		nickName = self.listBox.GetItemText(index)
		self.dialogActive = True
		# Translators: Message dialog to rename the nickname  of the selected folder.
		newKey = wx.GetTextFromUser(_("Enter a new nickname for %s") %nickName, title).strip().upper()
		if  newKey != '':
			if self.listBox.FindItem(0, newKey) == -1:
				config = ConfigObj(_ffIniFile, encoding = "utf-8")
				folders = config['Folders']
				try:
					nickName = nickName.decode('utf-8')
				except:
					pass
				path = folders[nickName]
				# update the dictionary.
				list = folders.items()
				list.append((newKey, path))
				list.remove((nickName, path))
				dictFolders = dict (list)
				try:
					newKey = newKey
				except:
					pass
				folders.rename(nickName, newKey)
				config.write()
				# update the list view.
				keys =folders.keys()
				keys.sort()
				newIndex = keys.index (newKey)
				self.updateFolders (folders, newIndex)
			else:
				gui.messageBox (_("There is already a folder with this nickname!"), title)
		self.dialogActive = False

	def onRemove (self, evt):
		# Removes the selected folder.
		evt.Skip()
		index=self.listBox.GetFocusedItem()
		nickName = self.listBox.GetItemText(index)
		self.dialogActive = True
		# Translators: Message dialog box to remove the selected folder.
		if gui.messageBox(_('Are you sure you want to remove %s?') %nickName, title, style=wx.ICON_QUESTION|wx.YES_NO) == wx.YES:
			config = ConfigObj(_ffIniFile, encoding = "utf-8")
			folders = config['Folders']
			folders.__delitem__(nickName)
			config.write()
			self.listBox.DeleteItem(index)
			if self.listBox.GetItemCount():
				self.listBox.Select(self.listBox.GetFocusedItem())
			else:
				self.setButtons()
		self.dialogActive = False

	def onKeyPress(self, evt):
		# Sets enter key  to open the folder and delete to remove it.
		evt.Skip()
		keycode = evt.GetKeyCode()
		if keycode == wx.WXK_RETURN and self.listBox.GetItemCount():
			self.onOpen(evt)
		elif keycode == wx.WXK_DELETE and self.listBox.GetItemCount():
			self.onRemove(evt)

	def setButtons(self):
		# Define which buttons are enabled and the check box is active.
		if self.listBox.GetItemCount():
			self.openButton.Enable()
			self.openButton.Show()
			self.renameButton.Enable()
			self.renameButton.Show()
			self.removeButton.Enable()
			self.removeButton.Show()
			try:
				if lastFocus.role == controlTypes.Role.EDITABLETEXT and not controlTypes.State.MULTILINE in lastFocus.states:
					self.pastButton.Enable()
					self.pastButton.Show()
				else:
					self.pastButton.Disable()
					self.pastButton.Hide()
			except AttributeError: 
				if lastFocus.role == controlTypes.ROLE_EDITABLETEXT and not controlTypes.STATE_MULTILINE in lastFocus.states:
					self.pastButton.Enable()
					self.pastButton.Show()
				else:
					self.pastButton.Disable()
					self.pastButton.Hide()
			self.chkAddress.	Enable()
			self.chkAddress.Show()	
		else:
			self.openButton.Disable()
			self.openButton.Hide()
			self.pastButton.Disable()
			self.pastButton.Hide()
			self.renameButton.Disable()
			self.renameButton.Hide()
			self.removeButton.Disable()
			self.removeButton.Hide()
			self.chkAddress.Disable()
			self.chkAddress.Hide()
		if newFolder is not None:
			self.addButton.Enable()
			self.addButton.Show()
		else:
			self.addButton.Disable()
			self.addButton.Hide()

	def onFocusLost(self,evt):
		# Close the add-on if the focus out of our window.
		obj = api.getForegroundObject()
		hasFocus= False
		for child in obj.children:
			hasFocus+= child.hasFocus
		if not hasFocus and not self.dialogActive:
			self.Hide()
			evt.Skip()

	def updateFolders(self, dictFolders, index):
		self.listBox.ClearAll()
		config = ConfigObj(_ffIniFile, encoding = "utf-8")
		# Translators: Title of the first column of the list view.
		self.listBox.InsertColumn(0, _('Nickname'))
		self.listBox.SetColumnWidth (0,250)
		if dictFolders == None:
			return
		try:
			if config['Options']['ShowPath'] == "1":
				# Translators: Title of the second column of the list view.
				self.listBox.InsertColumn(1, _('Address'))
				self.listBox.SetColumnWidth (1,500)
		except KeyError:
			pass
		keys = list(dictFolders.keys())
		keys.sort()
		cont = 0
		for item in keys:
			try:
				k = item.decode('utf-8')
			except:
				k = item
			try:
				v = dictFolders[item].decode('utf-8')
			except:
				v = dictFolders [item]
			self.listBox.Append ([k])
			try:
				if config['Options']['ShowPath'] == "1":
					self.listBox.SetStringItem (cont, 1, v)
			except KeyError:
				pass
			cont += 1
		self.listBox.Focus(index)
		self.listBox.Select(index)
		self.listBox.SetFocus()
