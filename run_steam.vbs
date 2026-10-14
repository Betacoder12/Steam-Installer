' Closes any running Steam, then launches it. No terminal window. Double-click to run.
Option Explicit

Dim sh, steamDir
Set sh = CreateObject("WScript.Shell")

' Edit this if your MSYS2 install or home folder is somewhere else:
steamDir = "C:\msys64\home\" & sh.ExpandEnvironmentStrings("%USERNAME%") & "\steam\steaminstall"

' Close Steam (0 = hidden window, True = wait for it to finish)
sh.Run "taskkill /F /T /IM steam.exe", 0, True
sh.Run "taskkill /F /T /IM steamwebhelper.exe", 0, True
WScript.Sleep 2000

' Launch Steam without admin prompts
sh.Environment("Process")("__COMPAT_LAYER") = "RUNASINVOKER"
sh.CurrentDirectory = steamDir
sh.Run """" & steamDir & "\steam.exe""", 1, False