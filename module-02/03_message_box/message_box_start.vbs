'--------------------------------------[ MsgBox Button & Icon Combinations ]--------------------------------------

' Simple OK + Cancel (no icon):
' 1 = vbOKCancel → 1 total
MsgBox "I am learning VB", 1, "1: OK + Cancel"


' Simple OK + Cancel with exclamation icon:
' 1 = vbOKCancel (buttons) + 48 = vbExclamation (icon) → 49 total
MsgBox "I am learning VB", 49, "49: OK + Cancel with Exclamation"



' Yes + No + Cancel with question icon:
' 3 = vbYesNoCancel (buttons) + 64 = vbQuestion (icon) → 67 total
MsgBox "I am learning VB", 67, "67: Yes + No + Cancel with Question"





'--------------------------------------[ MsgBox Button & Icon Reference Table ]--------------------------------------

' BUTTONS ONLY (without icons):
' Value | Constant             | Description
' ------|----------------------|------------------------
'   0   | vbOKOnly             | OK
'   1   | vbOKCancel           | OK, Cancel
'   2   | vbAbortRetryIgnore   | Abort, Retry, Ignore
'   3   | vbYesNoCancel        | Yes, No, Cancel
'   4   | vbYesNo              | Yes, No
'   5   | vbRetryCancel        | Retry, Cancel

' ICONS (can be combined with any button type):
' Value | Constant          | Icon Type
' ------|-------------------|---------------------
'  16   | vbCritical        | ❌ Stop / Error
'  32   | vbQuestion        | ❓ Question
'  48   | vbExclamation     | ⚠️ Warning
'  64   | vbInformation     | ℹ️ Info

' You combine Button + Icon using addition:
' Example: vbYesNoCancel (3) + vbQuestion (32) = 35

'--------------------------------------[ Button Types – Examples ]--------------------------------------

MsgBox "Example: OK only", 0, "Button: vbOKOnly (0)"
MsgBox "Example: OK + Cancel", 1, "Button: vbOKCancel (1)"
MsgBox "Example: Abort + Retry + Ignore", 2, "Button: vbAbortRetryIgnore (2)"
MsgBox "Example: Yes + No + Cancel", 3, "Button: vbYesNoCancel (3)"
MsgBox "Example: Yes + No", 4, "Button: vbYesNo (4)"
MsgBox "Example: Retry + Cancel", 5, "Button: vbRetryCancel (5)"

'--------------------------------------[ Icon Types – Examples with vbOKOnly ]--------------------------------------

MsgBox "Example: Critical icon", 16, "Icon: vbCritical (16)"
MsgBox "Example: Question icon", 32, "Icon: vbQuestion (32)"
MsgBox "Example: Exclamation icon", 48, "Icon: vbExclamation (48)"
MsgBox "Example: Information icon", 64, "Icon: vbInformation (64)"

'--------------------------------------[ Combined Examples – Buttons + Icons ]--------------------------------------

MsgBox "OK + Cancel + Warning", 1 + 48, "Combo: 49"
MsgBox "Yes + No + Cancel + Question", 3 + 32, "Combo: 35"
MsgBox "Retry + Cancel + Info", 5 + 64, "Combo: 69"
MsgBox "Abort + Retry + Ignore + Critical", 2 + 16, "Combo: 18"
