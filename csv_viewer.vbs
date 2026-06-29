' csv_viewer.vbs - CSV Viewer launcher (Windows)
'
' Usage:
'   - Double click: start the viewer (no file)
'   - Drop a CSV file onto this vbs icon: open that file
'   - Set as "Open with" handler for .csv
'
' Requires: Python 3 installed on Windows (tkinter is bundled).
'   For window drag and drop:  pip install tkinterdnd2
'
' No virtualenv needed. Uses the system Python.

Option Explicit

Dim fso, sh, scriptDir, pyScript, cmd, arg
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
pyScript  = scriptDir & "\csv_viewer.py"

' Clear env vars that break the Windows Python stdlib (often leaked from WSL)
Dim env
Set env = sh.Environment("PROCESS")
env.Remove "PYTHONHOME"
env.Remove "PYTHONPATH"
env.Remove "PYTHONSTARTUP"

If Not fso.FileExists(pyScript) Then
    MsgBox "csv_viewer.py not found:" & vbCrLf & pyScript, vbCritical, "CSV Viewer"
    WScript.Quit 1
End If

' Build command, quoting paths that may contain spaces.
' Use the 'pyw' launcher (windowed, no console). The bare python.exe on PATH
' may be a broken install, so the launcher is more reliable.
cmd = "pyw -3 """ & pyScript & """"

' Forward the dropped/associated file path (first one)
If WScript.Arguments.Count > 0 Then
    arg = WScript.Arguments(0)
    cmd = cmd & " """ & arg & """"
End If

' 0 = hidden console window, False = do not wait
On Error Resume Next
sh.Run cmd, 0, False
If Err.Number <> 0 Then
    ' Fallback to 'py -3' if 'pyw' is missing
    Err.Clear
    cmd = "py -3 """ & pyScript & """"
    If WScript.Arguments.Count > 0 Then cmd = cmd & " """ & WScript.Arguments(0) & """"
    sh.Run cmd, 1, False
    If Err.Number <> 0 Then
        MsgBox "Failed to start Python via the py launcher. Make sure Python 3 is installed.", vbCritical, "CSV Viewer"
    End If
End If
On Error GoTo 0
