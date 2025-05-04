Dim name, age, city

name = InputBox("Was ist dein Name?", "Willkommen")
age = InputBox("Wie alt bist du?", "Alter") ' akzeptiert negative Zahlen, da InputBox immer einen String liefert
city = InputBox("Wo wohnst du?", "Stadt")

MsgBox "Willkommen " & name & ", du bist " & age & " Jahre alt und kommst aus " & city & "."
