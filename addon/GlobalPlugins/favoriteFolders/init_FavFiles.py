# __init__.py
# Copyright (C) 2026 'Chai Chaimee based on the work of Rui Fontes
# Licensed under GNU General Public License. See COPYING.txt for details.

import os
import wx
import ui
import api
import gui
import socket
import globalVars
import winUser
import globalPluginHandler
import scriptHandler
import json
from ctypes import windll
from configobj import ConfigObj
from comtypes.client import CreateObject as COMCreate
import addonHandler

addonHandler.initTranslation()

# Use globalVars that has been imported
_FF_INI_FILE = os.path.abspath(os.path.join(globalVars.appArgs.configPath, "FavoriteFiles.ini"))
_FF_JSON_FILE = os.path.abspath(os.path.join(globalVars.appArgs.configPath, "FavoriteFiles.json"))
TITLE = _("Favorite files")


class FavoriteFilesGlobalPlugin(globalPluginHandler.GlobalPlugin):

    def __init__(self):
        super().__init__()
        self.dialog = None
        self.lastForeground = 0
        self._files = {}          # displayName: path
        self._pinned = set()      # Set of pinned item names
        self._showPath = False
        self._newFile = ""
        self._sortMode = "UPPERCASE"  # Default sort mode

    def _loadConfig(self):
        self._files = {}
        self._pinned = set()
        self._showPath = False
        self._sortMode = "UPPERCASE"

        # Load from INI file for backward compatibility
        if os.path.isfile(_FF_INI_FILE):
            config = ConfigObj(_FF_INI_FILE, encoding="utf-8")
            self._showPath = config.get("Options", {}).get("ShowPath", "0") == "1"

            # Load pinned items from INI
            pinned_section = config.get("Pinned", {})
            for name in pinned_section.keys():
                if name:
                    self._pinned.add(name)

            filesSection = config.get("Files", {})
            valid = {}
            for name, path in filesSection.items():
                if not path or not isinstance(path, str):
                    continue
                if path.startswith(r'\\'):
                    if self._isNetworkActive(path):
                        valid[name] = path
                else:
                    if os.path.isfile(path) and self._isLocalDriveValid(path):
                        valid[name] = path
            self._files = valid
            
            # Try to load sort mode from INI
            ini_sort_mode = config.get("Options", {}).get("SortMode", "")
            if ini_sort_mode in ["CUSTOM", "UPPERCASE", "LOWERCASE"]:
                self._sortMode = ini_sort_mode

        # Load from JSON file for extended settings
        if os.path.isfile(_FF_JSON_FILE):
            try:
                with open(_FF_JSON_FILE, 'r', encoding='utf-8') as f:
                    json_config = json.load(f)
                
                # Load sort mode from JSON
                if "sortMode" in json_config:
                    self._sortMode = json_config["sortMode"]
                
                # Load custom order from JSON
                if "customOrder" in json_config and "files" in json_config:
                    custom_order = json_config["customOrder"]
                    files_dict = json_config["files"]
                    
                    # Rebuild files dictionary in custom order
                    ordered_files = {}
                    for name in custom_order:
                        if name in files_dict:
                            path = files_dict[name]
                            # Validate file exists before adding
                            if path.startswith(r'\\'):
                                if self._isNetworkActive(path):
                                    ordered_files[name] = path
                            else:
                                if os.path.isfile(path) and self._isLocalDriveValid(path):
                                    ordered_files[name] = path
                    
                    # Add any remaining files not in custom order
                    for name, path in files_dict.items():
                        if name not in ordered_files:
                            if path.startswith(r'\\'):
                                if self._isNetworkActive(path):
                                    ordered_files[name] = path
                            else:
                                if os.path.isfile(path) and self._isLocalDriveValid(path):
                                    ordered_files[name] = path
                    
                    self._files = ordered_files
                
                # Load pinned items from JSON
                if "pinned" in json_config:
                    self._pinned = set(json_config["pinned"])
                
                # Load showPath from JSON
                if "showPath" in json_config:
                    self._showPath = json_config["showPath"]
                    
            except Exception as e:
                # If JSON loading fails, continue with INI data
                pass

    def _saveConfig(self):
        # Save to INI file for backward compatibility
        config = ConfigObj(_FF_INI_FILE, encoding="utf-8")
        config["Options"] = {
            "ShowPath": "1" if self._showPath else "0",
            "SortMode": self._sortMode
        }
        config["Files"] = self._files.copy()
        
        # Save pinned items to INI
        pinned_dict = {}
        for name in self._pinned:
            pinned_dict[name] = "1"
        config["Pinned"] = pinned_dict
        
        config.write()
        
        # Save to JSON file for extended settings
        json_config = {
            "sortMode": self._sortMode,
            "customOrder": list(self._files.keys()),  # Save current order
            "files": self._files.copy(),
            "pinned": list(self._pinned),
            "showPath": self._showPath
        }
        
        try:
            with open(_FF_JSON_FILE, 'w', encoding='utf-8') as f:
                json.dump(json_config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            # If JSON save fails, at least INI file is saved
            pass

    def _isLocalDriveValid(self, path):
        try:
            drive = path[:3]
            return windll.kernel32.GetDriveTypeW(drive) in (2, 3)
        except:
            return False

    def _isNetworkActive(self, path):
        try:
            if not path.startswith(r'\\'):
                return False
            host = path[2:].split('\\', 1)[0].strip('\\')
            ip = socket.gethostbyname(host)
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(1)
                s.connect((ip, 445))
            return os.path.isfile(path)
        except:
            return False

    def _getCurrentPathFromExplorer(self):
        fg = api.getForegroundObject()
        if not (fg.appModule and fg.appModule.appName == "explorer"):
            return None

        try:
            shell = COMCreate("Shell.Application")
            for window in shell.Windows():
                if window.hwnd == fg.windowHandle:
                    item = window.Document.FocusedItem
                    return item.Path if item else None
        except:
            return None

    def startFavoriteFiles(self):
        path = self._getCurrentPathFromExplorer()

        if path:
            if os.path.isfile(path):
                self._newFile = path
            elif os.path.isdir(path):
                ui.message(_("Folders cannot be added. Only files are allowed."))
                self._newFile = ""
            else:
                self._newFile = ""
        else:
            self._newFile = ""

        self._showDialog()

    def _showDialog(self):
        self._loadConfig()  # Load latest data every time

        if not self._files and not self._newFile:
            ui.message(_("You have no favorite files and are not in a Windows Explorer window to add one"))
            return

        # Check if new file is a duplicate
        if self._newFile and any(self._newFile.lower() == p.lower() for p in self._files.values()):
            self._newFile = ""

        try:
            self.lastForeground = api.getForegroundObject().windowHandle
        except:
            self.lastForeground = 0

        if self.dialog:
            self.dialog.Destroy()
            self.dialog = None

        self.dialog = FavoriteFilesDialog(gui.mainFrame, self)
        gui.mainFrame.prePopup()
        self.dialog.Show()
        self.dialog.CentreOnScreen()
        gui.mainFrame.postPopup()

    def terminate(self):
        if self.dialog:
            self.dialog.Destroy()


class FavoriteFilesDialog(wx.Dialog):
    def __init__(self, parent, plugin):
        super().__init__(parent, title=TITLE, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER | wx.MAXIMIZE_BOX)
        self.plugin = plugin
        self.currentSortMode = plugin._sortMode  # Use saved sort mode

        self._initUI()
        self._bindEvents()
        self._loadShowPath()
        self.updateFiles()
        self.setButtons()

    def _initUI(self):
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        self.listCtrl = wx.ListCtrl(self, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.BORDER_SUNKEN)
        self.listCtrl.InsertColumn(0, _("Display name"), width=280)
        self.listCtrl.InsertColumn(1, _("Path"), width=450)
        mainSizer.Add(self.listCtrl, 1, wx.EXPAND | wx.ALL, 10)

        optionsSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.chkShowPath = wx.CheckBox(self, label=_("&Show paths in the list"))
        optionsSizer.Add(self.chkShowPath, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 8)

        sortLabel = wx.StaticText(self, label=_("Sort order:"))
        optionsSizer.Add(sortLabel, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 8)

        self.sortCombo = wx.ComboBox(self, choices=[
            _("Custom order (drag and pin allowed)"),
            _("A-Z (uppercase)"),
            _("Z-A (lowercase)")
        ], style=wx.CB_READONLY)
        
        # Set selection based on saved sort mode
        if self.currentSortMode == "CUSTOM":
            self.sortCombo.SetSelection(0)
        elif self.currentSortMode == "UPPERCASE":
            self.sortCombo.SetSelection(1)
        else:  # LOWERCASE
            self.sortCombo.SetSelection(2)
            
        optionsSizer.Add(self.sortCombo, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 8)

        mainSizer.Add(optionsSizer, 0, wx.EXPAND | wx.ALL, 5)

        btnSizer = wx.GridSizer(2, 3, 10, 10)
        self.btnAdd = wx.Button(self, label=_("&Add"))
        self.btnOpen = wx.Button(self, label=_("&Open"))
        self.btnEdit = wx.Button(self, label=_("&Edit"))
        self.btnRemove = wx.Button(self, label=_("&Remove"))
        self.btnClose = wx.Button(self, wx.ID_CLOSE)

        for btn in (self.btnAdd, self.btnOpen, self.btnEdit, self.btnRemove, self.btnClose):
            btnSizer.Add(btn, 0, wx.EXPAND)

        mainSizer.Add(btnSizer, 0, wx.EXPAND | wx.ALL, 10)

        self.SetSizer(mainSizer)
        self.SetMinSize((800, 500))
        self.Fit()

        self.SetEscapeId(wx.ID_CLOSE)

    def _bindEvents(self):
        self.Bind(wx.EVT_CHECKBOX, self.onShowPathChanged, self.chkShowPath)
        self.Bind(wx.EVT_COMBOBOX, self.onSortChanged, self.sortCombo)

        self.Bind(wx.EVT_BUTTON, self.onAdd, self.btnAdd)
        self.Bind(wx.EVT_BUTTON, self.onOpen, self.btnOpen)
        self.Bind(wx.EVT_BUTTON, self.onEdit, self.btnEdit)
        self.Bind(wx.EVT_BUTTON, self.onRemove, self.btnRemove)
        self.Bind(wx.EVT_BUTTON, lambda e: self.Close(), self.btnClose)

        # Bind events exactly like the example
        self.listCtrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.onOpen)
        self.listCtrl.Bind(wx.EVT_KEY_DOWN, self.onKeyDown)
        self.Bind(wx.EVT_CLOSE, self.onClose)
        
        # Create and bind context menu exactly like the example
        self.contextMenu = wx.Menu()
        self.pinMenuItem = self.contextMenu.Append(wx.ID_ANY, _("&Pin to top"))
        self.moveUpMenuItem = self.contextMenu.Append(wx.ID_ANY, _("Move &up"))
        self.moveDownMenuItem = self.contextMenu.Append(wx.ID_ANY, _("Move &down"))
        self.editMenuItem = self.contextMenu.Append(wx.ID_ANY, _("&Edit"))
        self.removeMenuItem = self.contextMenu.Append(wx.ID_ANY, _("&Remove"))
        
        self.listCtrl.Bind(wx.EVT_CONTEXT_MENU, self.onContextMenu)
        self.Bind(wx.EVT_MENU, self.onTogglePin, self.pinMenuItem)
        self.Bind(wx.EVT_MENU, self.onMoveUp, self.moveUpMenuItem)
        self.Bind(wx.EVT_MENU, self.onMoveDown, self.moveDownMenuItem)
        self.Bind(wx.EVT_MENU, self.onEdit, self.editMenuItem)
        self.Bind(wx.EVT_MENU, self.onRemove, self.removeMenuItem)

    def _loadShowPath(self):
        self.chkShowPath.SetValue(self.plugin._showPath)

    def onShowPathChanged(self, evt):
        self.plugin._showPath = self.chkShowPath.GetValue()
        self.plugin._saveConfig()

    def onSortChanged(self, evt):
        idx = self.sortCombo.GetSelection()
        self.currentSortMode = ["CUSTOM", "UPPERCASE", "LOWERCASE"][idx]
        
        # Save sort mode to plugin
        self.plugin._sortMode = self.currentSortMode
        self.plugin._saveConfig()
        
        self.updateFiles()

    def onContextMenu(self, evt):
        # Get mouse position from event
        pos = evt.GetPosition()
        
        # For context menu on list control, we need to check if we clicked on an item
        # But in the example, they just show the menu at the click position
        
        # First, check if we have any items
        if self.listCtrl.GetItemCount() == 0:
            evt.Skip()
            return
            
        # Get the item at the click position
        pos = self.listCtrl.ScreenToClient(pos)
        idx, flags = self.listCtrl.HitTest(pos)
        
        # If clicked on an item, select it
        if idx != -1:
            self.listCtrl.Select(idx)
            self.listCtrl.Focus(idx)
            self.listCtrl.EnsureVisible(idx)
        
        # Enable/disable menu items based on context
        hasSelection = self.listCtrl.GetFirstSelected() != -1
        
        # Enable/disable move up/down based on selection and sort mode
        if hasSelection and self.currentSortMode == "CUSTOM":
            selIdx = self.listCtrl.GetFirstSelected()
            self.moveUpMenuItem.Enable(selIdx > 0)
            self.moveDownMenuItem.Enable(selIdx < self.listCtrl.GetItemCount() - 1)
        else:
            self.moveUpMenuItem.Enable(False)
            self.moveDownMenuItem.Enable(False)
            
        # Enable other menu items if we have a selection
        self.editMenuItem.Enable(hasSelection)
        self.removeMenuItem.Enable(hasSelection)
        self.pinMenuItem.Enable(hasSelection)
        
        # Update pin/unpin menu text based on whether item is pinned
        if hasSelection:
            selIdx = self.listCtrl.GetFirstSelected()
            name = self.listCtrl.GetItemText(selIdx, 0)
            if name in self.plugin._pinned:
                self.pinMenuItem.SetItemLabel(_("&Unpin"))
            else:
                self.pinMenuItem.SetItemLabel(_("&Pin to top"))
        
        # Show the context menu at the mouse position
        self.listCtrl.PopupMenu(self.contextMenu, evt.GetPosition())

    def onTogglePin(self, evt):
        idx = self.listCtrl.GetFirstSelected()
        if idx == -1:
            return
            
        name = self.listCtrl.GetItemText(idx, 0)
        path = self.plugin._files.get(name)
        if not path:
            return

        if name in self.plugin._pinned:
            # Unpin - remove from pinned set
            self.plugin._pinned.remove(name)
            
            # In uppercase/lowercase mode, just remove pin and resort
            if self.currentSortMode in ["UPPERCASE", "LOWERCASE"]:
                # Just remove from pinned set and resort
                self.plugin._saveConfig()
                self.updateFiles()
                ui.message(_("Unpinned {}").format(name))
            else:
                # In CUSTOM mode, move to bottom
                # First, remove from current position
                del self.plugin._files[name]
                # Then add to end
                self.plugin._files[name] = path
                self.plugin._saveConfig()
                self.updateFiles()
                ui.message(_("Unpinned {}").format(name))
        else:
            # Pin to top
            self.plugin._pinned.add(name)
            
            # In uppercase/lowercase mode, we need to create a hybrid sort:
            # 1. Pinned items at the top (in alphabetical order)
            # 2. Unpinned items below (in alphabetical order)
            if self.currentSortMode in ["UPPERCASE", "LOWERCASE"]:
                # Get all items
                all_items = list(self.plugin._files.items())
                
                # Separate pinned and unpinned
                pinned_items = []
                unpinned_items = []
                
                for item_name, item_path in all_items:
                    if item_name in self.plugin._pinned:
                        pinned_items.append((item_name, item_path))
                    else:
                        unpinned_items.append((item_name, item_path))
                
                # Sort each group based on current sort mode
                if self.currentSortMode == "UPPERCASE":
                    pinned_items.sort(key=lambda x: x[0].upper())
                    unpinned_items.sort(key=lambda x: x[0].upper())
                else:  # LOWERCASE
                    pinned_items.sort(key=lambda x: x[0].lower(), reverse=True)
                    unpinned_items.sort(key=lambda x: x[0].lower(), reverse=True)
                
                # Combine: pinned first, then unpinned
                sorted_items = pinned_items + unpinned_items
                
                # Update dictionary with new order
                self.plugin._files = dict(sorted_items)
                
                self.plugin._saveConfig()
                # Find and select the pinned item
                for i, (item_name, _) in enumerate(sorted_items):
                    if item_name == name:
                        self.updateFiles(i)
                        break
                ui.message(_("Pinned {} to top").format(name))
            else:
                # Already in CUSTOM mode, move pinned item to top
                # Remove from current position
                del self.plugin._files[name]
                
                # Create new dictionary with pinned item first
                new_files = {name: path}
                new_files.update(self.plugin._files)
                self.plugin._files = new_files
                
                self.plugin._saveConfig()
                # Update and select the pinned item (should be at position 0)
                self.updateFiles(0)
                ui.message(_("Pinned {} to top").format(name))

    def onMoveUp(self, evt):
        idx = self.listCtrl.GetFirstSelected()
        if idx <= 0 or self.currentSortMode != "CUSTOM":
            return
            
        items = list(self.plugin._files.items())
        items[idx-1], items[idx] = items[idx], items[idx-1]
        self.plugin._files = dict(items)
        self.plugin._saveConfig()
        self.updateFiles(idx-1)

    def onMoveDown(self, evt):
        idx = self.listCtrl.GetFirstSelected()
        if idx >= len(self.plugin._files) - 1 or self.currentSortMode != "CUSTOM":
            return
            
        items = list(self.plugin._files.items())
        items[idx], items[idx+1] = items[idx+1], items[idx]
        self.plugin._files = dict(items)
        self.plugin._saveConfig()
        self.updateFiles(idx+1)

    def onAdd(self, evt):
        if not self.plugin._newFile:
            ui.message(_("No file selected to add"))
            return

        default = os.path.basename(self.plugin._newFile)
        dlg = wx.TextEntryDialog(self, _("Enter display name for the file"), TITLE, default)
        if dlg.ShowModal() == wx.ID_OK:
            name = dlg.GetValue().strip()
            if not name:
                gui.messageBox(_("Display name cannot be empty"), TITLE, wx.OK | wx.ICON_WARNING)
                dlg.Destroy()
                return
            if name in self.plugin._files:
                gui.messageBox(_("This display name already exists!"), TITLE, wx.OK | wx.ICON_WARNING)
                dlg.Destroy()
                return

            self.plugin._files[name] = self.plugin._newFile
            self.plugin._saveConfig()
            self.updateFiles(len(self.plugin._files) - 1)
            self.listCtrl.Select(len(self.plugin._files) - 1)
            self.listCtrl.Focus(len(self.plugin._files) - 1)
            ui.message(_("Added: {}").format(name))
        dlg.Destroy()
        self.setButtons()

    def onOpen(self, evt):
        idx = self.listCtrl.GetFirstSelected()
        if idx == -1:
            return
        name = self.listCtrl.GetItemText(idx, 0)
        path = self.plugin._files.get(name)
        if path and os.path.isfile(path):
            self.Hide()
            os.startfile(path)
            self.Close()
        else:
            ui.message(_("File not found"))

    def onEdit(self, evt):
        idx = self.listCtrl.GetFirstSelected()
        if idx == -1:
            return
        oldName = self.listCtrl.GetItemText(idx, 0)
        path = self.plugin._files.get(oldName)
        if not path:
            return

        dlg = wx.TextEntryDialog(self, _("Enter new display name"), TITLE, oldName)
        if dlg.ShowModal() == wx.ID_OK:
            newName = dlg.GetValue().strip()
            if not newName or newName == oldName:
                dlg.Destroy()
                return
            if newName in self.plugin._files:
                gui.messageBox(_("This display name already exists!"), TITLE, wx.OK | wx.ICON_WARNING)
                dlg.Destroy()
                return

            # Update pinned set if needed
            if oldName in self.plugin._pinned:
                self.plugin._pinned.remove(oldName)
                self.plugin._pinned.add(newName)
            
            # Update position in dictionary to maintain order
            items = list(self.plugin._files.items())
            for i, (n, p) in enumerate(items):
                if n == oldName:
                    items[i] = (newName, path)
                    break
            
            self.plugin._files = dict(items)
            self.plugin._saveConfig()
            self.updateFiles(idx)
            self.listCtrl.Select(idx)
            self.listCtrl.Focus(idx)
            ui.message(_("Display name updated to {}").format(newName))
        dlg.Destroy()

    def onRemove(self, evt):
        idx = self.listCtrl.GetFirstSelected()
        if idx == -1:
            return
        name = self.listCtrl.GetItemText(idx, 0)

        if gui.messageBox(
            _("Are you sure you want to remove {}?").format(name),
            TITLE,
            wx.YES_NO | wx.ICON_QUESTION
        ) == wx.YES:
            # Remove from pinned set if present
            if name in self.plugin._pinned:
                self.plugin._pinned.remove(name)
            
            del self.plugin._files[name]
            self.plugin._saveConfig()
            self.updateFiles(max(0, idx - 1))
            self.setButtons()
            ui.message(_("Removed {}").format(name))

    def onKeyDown(self, evt):
        key = evt.GetKeyCode()
        if key == wx.WXK_RETURN:
            self.onOpen(evt)
        elif key == wx.WXK_DELETE:
            self.onRemove(evt)
        else:
            evt.Skip()

    def setButtons(self):
        hasItems = self.listCtrl.GetItemCount() > 0
        hasSelection = self.listCtrl.GetFirstSelected() != -1
        self.btnOpen.Enable(hasSelection)
        self.btnEdit.Enable(hasSelection)
        self.btnRemove.Enable(hasSelection)
        self.btnAdd.Enable(bool(self.plugin._newFile))
        self.chkShowPath.Enable(hasItems)

    def updateFiles(self, selectIdx=0):
        self.listCtrl.DeleteAllItems()
        if not self.plugin._files:
            self.setButtons()
            return

        show = self.chkShowPath.GetValue()

        if self.currentSortMode == "CUSTOM":
            # In CUSTOM mode, show items in their current order
            items = list(self.plugin._files.items())
        elif self.currentSortMode == "UPPERCASE":
            # In UPPERCASE mode, sort with pinned items first, then unpinned items
            # Both groups sorted A-Z (case-insensitive)
            all_items = list(self.plugin._files.items())
            pinned_items = []
            unpinned_items = []
            
            for item_name, item_path in all_items:
                if item_name in self.plugin._pinned:
                    pinned_items.append((item_name, item_path))
                else:
                    unpinned_items.append((item_name, item_path))
            
            # Sort each group A-Z (case-insensitive)
            pinned_items.sort(key=lambda x: x[0].upper())
            unpinned_items.sort(key=lambda x: x[0].upper())
            
            # Combine: pinned items first, then unpinned items
            items = pinned_items + unpinned_items
        else:  # LOWERCASE
            # In LOWERCASE mode, sort with pinned items first, then unpinned items
            # Both groups sorted Z-A (case-insensitive)
            all_items = list(self.plugin._files.items())
            pinned_items = []
            unpinned_items = []
            
            for item_name, item_path in all_items:
                if item_name in self.plugin._pinned:
                    pinned_items.append((item_name, item_path))
                else:
                    unpinned_items.append((item_name, item_path))
            
            # Sort each group Z-A (case-insensitive)
            pinned_items.sort(key=lambda x: x[0].lower(), reverse=True)
            unpinned_items.sort(key=lambda x: x[0].lower(), reverse=True)
            
            # Combine: pinned items first, then unpinned items
            items = pinned_items + unpinned_items

        for i, (name, path) in enumerate(items):
            self.listCtrl.InsertItem(i, name)
            if show:
                self.listCtrl.SetItem(i, 1, path)

        if 0 <= selectIdx < self.listCtrl.GetItemCount():
            self.listCtrl.Select(selectIdx)
            self.listCtrl.Focus(selectIdx)
            self.listCtrl.EnsureVisible(selectIdx)

        self.setButtons()

    def onClose(self, evt):
        # Save any pending changes before closing
        self.plugin._saveConfig()
        self.Destroy()
        evt.Skip()


GlobalPlugin = FavoriteFilesGlobalPlugin
