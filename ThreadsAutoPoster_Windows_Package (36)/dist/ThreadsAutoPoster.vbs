' Threads AutoPoster - Smart App Control Bypass Launcher (100% ASCII-Safe)
Option Explicit
Dim WshShell, fso, currentDir, foundPy, scriptPath, subPath, ans, setupBat, isTemp
Dim msgZip, titleZip, msgNoScript, titleNoScript, msgNoPy, titleNoPy

Function JText(s)
    Dim res, parts, i, val
    parts = Split(s, ",")
    res = ""
    For i = 0 To UBound(parts)
        val = Trim(parts(i))
        If val <> "" Then res = res & ChrW(CLng(val))
    Next
    JText = res
End Function

Function CheckPy(p)
    CheckPy = False
    If InStr(1, p, "WindowsApps", 1) > 0 Then Exit Function
    If InStr(p, "\") > 0 Then
        If Not fso.FileExists(p) Then Exit Function
    End If
    On Error Resume Next
    Dim ret
    ret = WshShell.Run("""" & p & """ -c ""import sys""", 0, True)
    If Err.Number = 0 And ret = 0 Then
        CheckPy = True
    End If
    On Error GoTo 0
End Function

Function FindPy()
    FindPy = ""
    ' Try registry first
    Dim regRoots, ver, root, base
    regRoots = Array("HKCU\Software\Python\PythonCore\", "HKLM\Software\Python\PythonCore\")
    For Each root In regRoots
        For Each ver In Array("3.13", "3.12", "3.11", "3.10", "3.9", "3.8")
            On Error Resume Next
            base = WshShell.RegRead(root & ver & "\InstallPath\")
            If Err.Number = 0 And base <> "" Then
                If fso.FileExists(base & "pythonw.exe") Then
                    FindPy = base & "pythonw.exe"
                    On Error GoTo 0
                    Exit Function
                ElseIf fso.FileExists(base & "python.exe") Then
                    FindPy = base & "python.exe"
                    On Error GoTo 0
                    Exit Function
                End If
            End If
            On Error GoTo 0
        Next
    Next

    ' Try PATH and standard directories
    Dim cand, candidates
    candidates = Array(_
        "pythonw.exe", _
        "python.exe", _
        "py.exe", _
        WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python312\pythonw.exe", _
        WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python313\pythonw.exe", _
        WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python311\pythonw.exe", _
        WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python310\pythonw.exe", _
        "C:\Program Files\Python312\pythonw.exe", _
        "C:\Program Files\Python313\pythonw.exe", _
        "C:\Program Files\Python311\pythonw.exe", _
        "C:\Program Files\Python310\pythonw.exe", _
        WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python312\python.exe", _
        WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python313\python.exe", _
        "C:\Program Files\Python312\python.exe", _
        "C:\Program Files\Python313\python.exe" _
    )
    For Each cand In candidates
        If CheckPy(cand) Then
            FindPy = cand
            Exit Function
        End If
    Next
End Function

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
currentDir = fso.GetParentFolderName(WScript.ScriptFullName)

' 1. Check if running inside temporary unzipped folder
isTemp = False
If InStr(1, currentDir, "\Temp", 1) > 0 Or InStr(1, currentDir, "Temp1_", 1) > 0 Then
    isTemp = True
End If

' 2. Locate threads_gui_app.py
scriptPath = currentDir & "\threads_gui_app.py"
If Not fso.FileExists(scriptPath) Then
    subPath = currentDir & "\ThreadsAutoPoster\threads_gui_app.py"
    If fso.FileExists(subPath) Then
        scriptPath = subPath
        currentDir = currentDir & "\ThreadsAutoPoster"
    Else
        If isTemp Then
            msgZip = JText("12304,90,73,80,12501,12449,12452,12523,12398,12300,12377,12409,12390,23637,38283,12301,12364,24517,35201,12391,12377,12305,10,10,12480,12454,12531,12525,12540,12489,12375,12383,90,73,80,12501,12449,12452,12523,12434,38283,12356,12383,12414,12414,30452,25509,23455,34892,12373,12428,12390,12356,12414,12377,12290,10,87,105,110,100,111,119,115,12398,20181,27096,19978,12289,90,73,80,20869,12363,12425,30452,25509,23455,34892,12377,12427,12392,20182,12398,24517,38920,12501,12449,12452,12523,12364,35501,12415,36796,12417,12414,12379,12435,12290,10,10,12304,35299,27770,25163,38918,12305,10,49,46,32,12480,12454,12531,12525,12540,12489,12375,12383,90,73,80,12501,12449,12452,12523,12434,21491,12463,12522,12483,12463,10,50,46,32,12300,12377,12409,12390,23637,38283,12301,12434,12463,12522,12483,12463,12375,12390,35299,20941,10,51,46,32,23637,38283,12373,12428,12383,26032,12375,12356,12501,12457,12523,12480,20869,12398,12300,84,104,114,101,97,100,115,65,117,116,111,80,111,115,116,101,114,46,118,98,115,12301,12414,12383,12399,12300,115,116,97,114,116,95,103,117,105,46,98,97,116,12301,12434,38283,12356,12390,12367,12384,12373,12356,12290")
            titleZip = JText("84,104,114,101,97,100,115,32,65,117,116,111,80,111,115,116,101,114,32,45,32,35299,20941,12364,24517,35201,12391,12377")
            MsgBox msgZip, 48, titleZip
        Else
            msgNoScript = JText("116,104,114,101,97,100,115,95,103,117,105,95,97,112,112,46,112,121,32,12364,35211,12388,12363,12426,12414,12379,12435,12290,10,10,35299,20941,12373,12428,12383,12501,12457,12523,12480,19968,24335,65288,69,120,99,101,108,12420,12473,12463,12522,12503,12488,12364,20840,12390,20837,12387,12390,12356,12427,12501,12457,12523,12480,65289,12398,20013,12391,23455,34892,12375,12390,12367,12384,12373,12356,12290")
            titleNoScript = JText("84,104,114,101,97,100,115,32,65,117,116,111,80,111,115,116,101,114,32,45,32,12456,12521,12540")
            MsgBox msgNoScript, 16, titleNoScript
        End If
        WScript.Quit 1
    End If
End If

' 3. Find Python and run
foundPy = FindPy()

If foundPy = "" Then
    msgNoPy = JText("84,104,114,101,97,100,115,32,65,117,116,111,80,111,115,116,101,114,12398,23455,34892,12395,24517,35201,12394,80,121,116,104,111,110,12364,35211,12388,12363,12426,12414,12379,12435,12391,12375,12383,12290,10,10,12300,48,95,20182,12398,80,67,12391,12418,12377,12368,21205,12367,29872,22659,20316,25104,46,98,97,116,12301,12434,23455,34892,12375,12390,33258,21205,12475,12483,12488,12488,12450,12483,12503,12434,34892,12356,12414,12377,12363,65311")
    titleNoPy = JText("84,104,114,101,97,100,115,32,65,117,116,111,80,111,115,116,101,114,32,45,32,21021,26399,35373,23450")
    ans = MsgBox(msgNoPy, 36, titleNoPy)
    If ans = 6 Then ' vbYes
        setupBat = currentDir & "\setup_env.bat"
        If Not fso.FileExists(setupBat) Then setupBat = currentDir & "\start_gui.bat"
        If fso.FileExists(setupBat) Then
            WshShell.CurrentDirectory = currentDir
            WshShell.Run """" & setupBat & """", 1, False
        Else
            WshShell.Run "https://www.python.org/downloads/"
        End If
    End If
    WScript.Quit 1
End If

' 4. Launch GUI Application
WshShell.CurrentDirectory = currentDir
Dim winStyle, pyDir
If InStr(1, foundPy, "pythonw.exe", 1) = 0 And InStr(1, foundPy, "pyw.exe", 1) = 0 Then
    pyDir = fso.GetParentFolderName(foundPy)
    If pyDir <> "" And fso.FileExists(pyDir & "\pythonw.exe") Then
        foundPy = pyDir & "\pythonw.exe"
    ElseIf pyDir <> "" And fso.FileExists(pyDir & "\pyw.exe") Then
        foundPy = pyDir & "\pyw.exe"
    End If
End If
winStyle = 0
WshShell.Run """" & foundPy & """ """ & scriptPath & """", winStyle, False
