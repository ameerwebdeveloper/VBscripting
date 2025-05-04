'--------------------------------------[ Using InputBox to Get User Input ]--------------------------------------

Const App_Title = "→ Learning Input"

Dim numberOne, numberTwo, sum, name

' Get user input
name = InputBox("Deinen Namen eingeben:", "Willkommen")
numberOne = InputBox("Erste Zahl eingeben:", App_Title)
numberTwo = InputBox("Zweite Zahl eingeben:", App_Title)

' Convert input from String to Integer using CInt
' InputBox returns strings, so we must cast them to numbers to perform addition

sum = CInt(numberOne) + CInt(numberTwo)

' Output results
MsgBox "Hallo " & name
MsgBox "Das Ergebnis ist " & sum

'--------------------------------------[ Hinweis: Fehlerbehandlung ]--------------------------------------
' Wenn eines der Eingabefelder leer bleibt oder ungültig ist (z. B. Buchstaben statt Zahlen),
' führt CInt zu einem Laufzeitfehler:
' → Microsoft VBScript runtime error: Type mismatch: 'CInt'
' Lösung: Vor CInt prüfen, ob Eingaben gültig und nicht leer sind 
