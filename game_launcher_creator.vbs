' Creates a launcher .vbs for every Steam game, in the same folder as this file.
Option Explicit

Dim sh, fso, steamDir, commonDir, fo, exes, exeRel, outDir
Dim created, skipped, summary, re

Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set re = CreateObject("VBScript.RegExp")
re.Global = True

steamDir = "C:\msys64\home\" & sh.ExpandEnvironmentStrings("%USERNAME%") & "\steam\steaminstall"
commonDir = steamDir & "\steamapps\common"
outDir = fso.GetParentFolderName(WScript.ScriptFullName)

If Not fso.FolderExists(commonDir) Then
    MsgBox "Couldn't find:" & vbCrLf & commonDir, 16, "Game Launcher Creator"
    WScript.Quit
End If

created = ""
skipped = ""

For Each fo In fso.GetFolder(commonDir).SubFolders
    ' Skip Steam's shared redistributables
    If LCase(fo.Name) <> "steamworks shared" Then
        Set exes = CreateObject("Scripting.Dictionary")
        FindExes fo.Path, fo.Path

        If exes.Count = 0 Then
            skipped = skipped & "  - " & fo.Name & vbCrLf
        Else
            exeRel = PickExe(fo.Name)
            WriteLauncher fo.Name, exeRel
            created = created & "  - " & fo.Name & "  (" & exeRel & ")" & vbCrLf
        End If
    End If
Next

summary = ""
If created <> "" Then summary = "Created launchers for:" & vbCrLf & created
If skipped <> "" Then summary = summary & vbCrLf & "Skipped (no usable exe found):" & vbCrLf & skipped
If summary = "" Then summary = "No games found in:" & vbCrLf & commonDir

MsgBox summary, 64, "Game Launcher Creator"

' ---- Helpers ----
Sub WriteLauncher(gameFolder, exe)
    Dim f, safeName, outPath
    re.Pattern = "[\\/:*?""<>|]"
    safeName = re.Replace(gameFolder, "")
    outPath = outDir & "\Launch " & safeName & ".vbs"

    Set f = fso.CreateTextFile(outPath, True)
    f.WriteLine "' Launches Steam, waits for you, then launches " & gameFolder & ". Double-click to run."
    f.WriteLine "Option Explicit"
    f.WriteLine ""
    f.WriteLine "Dim sh, steamDir, gameDir, gameExe"
    f.WriteLine "Set sh = CreateObject(""WScript.Shell"")"
    f.WriteLine ""
    f.WriteLine "steamDir = ""C:\msys64\home\"" & sh.ExpandEnvironmentStrings(""%USERNAME%"") & ""\steam\steaminstall"""
    f.WriteLine "gameDir = steamDir & ""\steamapps\common\" & gameFolder & """"
    f.WriteLine "gameExe = gameDir & ""\" & exe & """"
    f.WriteLine ""
    f.WriteLine "sh.Run ""taskkill /F /T /IM steam.exe"", 0, True"
    f.WriteLine "sh.Run ""taskkill /F /T /IM steamwebhelper.exe"", 0, True"
    f.WriteLine "WScript.Sleep 2000"
    f.WriteLine ""
    f.WriteLine "sh.Environment(""Process"")(""__COMPAT_LAYER"") = ""RUNASINVOKER"""
    f.WriteLine "sh.CurrentDirectory = steamDir"
    f.WriteLine "sh.Run Chr(34) & steamDir & ""\steam.exe"" & Chr(34), 1, False"
    f.WriteLine ""
    f.WriteLine "MsgBox ""Dismiss the Steam pop-ups (Ignore / Cancel) and wait for Steam to finish loading, then click OK to launch the game."", 64, ""Game Launcher"""
    f.WriteLine ""
    f.WriteLine "sh.CurrentDirectory = gameDir"
    f.WriteLine "sh.Run Chr(34) & gameExe & Chr(34), 1, False"
    f.Close
End Sub

' Prefers an exe whose name matches the game folder; otherwise the first one found
Function PickExe(gameFolder)
    Dim i, folderKey, exeKey
    re.Pattern = "[^a-z0-9]"
    folderKey = re.Replace(LCase(gameFolder), "")
    For i = 1 To exes.Count
        exeKey = re.Replace(LCase(fso.GetBaseName(exes(i))), "")
        If exeKey = folderKey Then
            PickExe = exes(i)
            Exit Function
        End If
    Next
    PickExe = exes(1)
End Function

Sub FindExes(path, root)
    Dim fld, file, subFld
    Set fld = fso.GetFolder(path)
    For Each file In fld.Files
        If LCase(fso.GetExtensionName(file.Name)) = "exe" Then
            If Not IsJunk(file.Name) Then
                exes.Add exes.Count + 1, Mid(file.Path, Len(root) + 2)
            End If
        End If
    Next
    For Each subFld In fld.SubFolders
        FindExes subFld.Path, root
    Next
End Sub

Function IsJunk(fileName)
    Dim junk, j, n
    junk = Array("unins", "crashhandler", "crashreport", "vcredist", "vc_redist", "dxsetup", "dotnet", "ue4prereq", "uninstall", "setup")
    n = LCase(fileName)
    IsJunk = False
    For Each j In junk
        If InStr(n, j) > 0 Then
            IsJunk = True
            Exit Function
        End If
    Next
End Function