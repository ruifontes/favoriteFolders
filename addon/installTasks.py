#-*- coding: utf-8 -*-
# Part of favoriteFolders add-on for NVDA.
# written by Rui Fontes <rui.fontes@tiflotecnia.com>, Ângelo Abrantes and Abel Passos Júnior

import os
import globalVars
import addonHandler

def onInstall():
	configFilePath = os.path.abspath(os.path.join(globalVars.appArgs.configPath, "addons", "favoriteFolders", "globalPlugins", "favoriteFolders", "FavoriteFolders.ini"))
	if os.path.isfile(configFilePath):
		os.rename(configFilePath, os.path.abspath(os.path.join(globalVars.appArgs.configPath, "FavoriteFolders.ini")))
