' Downloads and runs the Steam installer through MSYS2. Double-click to run.
Option Explicit

Dim sh, cmd
Set sh = CreateObject("WScript.Shell")

Const MSYS_ROOT = "C:\msys64"

cmd = "cd ~ && " & _
      "mkdir -p steam && " & _
      "cd steam && " & _
      "mkdir -p steaminstall && " & _
      "wget https://steamcdn-a.akamaihd.net/client/installer/SteamSetup.exe && " & _
      "export __COMPAT_LAYER=RUNASINVOKER && " & _
      "./SteamSetup.exe"

sh.Run """" & MSYS_ROOT & "\usr\bin\env.exe"" MSYSTEM=UCRT64 /usr/bin/bash -lc """ & cmd & """", 1, True