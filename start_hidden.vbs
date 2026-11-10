' Starts the Procurement Test Runner in the background (no console window, no browser popup).
' Open http://127.0.0.1:8765 whenever you want to use it. Log: test_runner.log
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = dir
sh.Environment("Process")("TESTER_NO_BROWSER") = "1"
sh.Run "pythonw """ & dir & "\procurement_tester.py""", 0, False
