'--------------------------------------[ Simple Calculation Example ]--------------------------------------

Dim num1, num2, sum

num1 = 10 
num2 = 20
sum = num1 + num2

MsgBox sum

'--------------------------------------[ Better Naming for Variables ]--------------------------------------

Dim transactionOne, transactionTwo, totalTransfer

transactionOne = 400
transactionTwo = 500
totalTransfer = transactionOne + transactionTwo

WScript.Echo "Total transaction is " & totalTransfer

'--------------------------------------[ Constant Declaration ]--------------------------------------

Const bonus = 5
' bonus = 50  -> Microsoft VBScript runtime error: Illegal assignment: 'bonus'

totalTransfer = transactionOne + transactionTwo + bonus

WScript.Echo "Total transaction with bonus: " & totalTransfer

'--------------------------------------[ String Concatenation with & ]--------------------------------------

Dim ticketNr, strCustomerName, strTrouble

ticketNr = 1824
strCustomerName = "Max Mustermann"
strTrouble = "hat kein Internet"

WScript.Echo "Der Kunde mit Ticket-Nr. " & ticketNr & " und Name " & strCustomerName & " hat das Problem: " & strTrouble

'--------------------------------------[ Line Continuation with &_ ]--------------------------------------

WScript.Echo "Der Kunde mit Ticket-Nr. " & ticketNr & " und Name " & strCustomerName & _
    " hat das Problem: " & strTrouble

'--------------------------------------[ Convention: Prefix for String Types ]--------------------------------------

strName = "Ameer"
age = 30

WScript.Echo strName & " is " & age
