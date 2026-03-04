Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Dir = FSO.GetParentFolderName(WScript.ScriptFullName)
WshShell.Environment("Process")("VIRTUAL_ENV") = ""
WshShell.Environment("Process")("PYTHONHOME") = ""
WshShell.Environment("Process")("PYTHONPATH") = ""
WshShell.CurrentDirectory = WshShell.Environment("Process")("TEMP")
WshShell.Run "cmd /c cd /d " & Dir & " && py -3 -m pip install -r requirements.txt -q 2>nul && py -3 desktop_app.py", 0, False
