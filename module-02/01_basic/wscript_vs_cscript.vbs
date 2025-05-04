' WScript.EXE -> Windows Apps  
' CScript.EXE -> Console Apps  
' WScript -> is an object, that's why WScript.Echo  

MsgBox "1_Hello there"
WScript.Echo "2_Hello world! From WScript object"

' Sleep for 5 seconds
WScript.Sleep 5000

WScript.Echo "3_Hello world! after sleeping"

WScript.Quit
WScript.Echo "4_Hello world! after Quit"
