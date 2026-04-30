Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
pythonw = folder & "\.venv\Scripts\pythonw.exe"
python = folder & "\.venv\Scripts\python.exe"
cfg = folder & "\.venv\pyvenv.cfg"
gui = folder & "\gui.py"
launcher = folder & "\START_GUI.cmd"

Function VenvPythonMissing()
    VenvPythonMissing = False
    If Not fso.FileExists(cfg) Then Exit Function

    Set cfgFile = fso.OpenTextFile(cfg, 1, False)
    Do Until cfgFile.AtEndOfStream
        line = cfgFile.ReadLine
        If LCase(Left(line, 13)) = "executable = " Then
            target = Trim(Mid(line, 14))
            If target <> "" And Not fso.FileExists(target) Then
                VenvPythonMissing = True
            End If
            Exit Do
        End If
    Loop
    cfgFile.Close
End Function

Sub RunLauncher()
    If fso.FileExists(launcher) Then
        shell.CurrentDirectory = folder
        shell.Run """" & launcher & """", 1, False
    Else
        MsgBox "START_GUI.cmd was not found." & vbCrLf & launcher, 16, "Launch error"
    End If
End Sub

If Not fso.FileExists(pythonw) Then
    RunLauncher
    WScript.Quit 0
End If
If Not fso.FileExists(python) Then
    RunLauncher
    WScript.Quit 0
End If
If Not fso.FileExists(gui) Then
    MsgBox "gui.py was not found." & vbCrLf & gui, 16, "Launch error"
    WScript.Quit 1
End If
If VenvPythonMissing() Then
    RunLauncher
    WScript.Quit 0
End If

check = """" & python & """ -c ""import tkinter, selenium, docx, webdriver_manager"""
exitCode = shell.Run(check, 0, True)
If exitCode <> 0 Then
    RunLauncher
    WScript.Quit 0
End If

cmd = """" & pythonw & """ """ & gui & """"
shell.CurrentDirectory = folder
shell.Run cmd, 0, False
