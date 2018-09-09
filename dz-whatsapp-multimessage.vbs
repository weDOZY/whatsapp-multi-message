x=msgbox ("Are ready to proceed?" ,4+48, "Important Notice. Dozy ©")


x=msgbox ("Do you have Internet Connection Working?" ,4+48, "Important Notice. Dozy")

' InputBoxes
Contact = InputBox("Which Contact Do You Want To Send?", "WhatsApp MultiMessage - Dozy.in")
Message = InputBox("What Is The Message?","WhatsApp")
T = InputBox("How Many Times Needs It To Be Send?","WhatsApp")
If MsgBox("You've Filled It In Correctely?", 1024 + vbSystemModal, "WhatsApp MultiMessage - Dozy.in") = vbOk Then

' Go To WhatsApp
Set WshShell = WScript.CreateObject("WScript.Shell")
Return = WshShell.Run("https://web.whatsapp.com/", 1)

' Loading Time

If MsgBox("Is WhatsApp Loaded?" & vbNewLine & vbNewLine & "Press No To Cancel", vbYesNo + vbQuestion + vbSystemModal, "WhatsApp MultiMessage - Dozy.in") = vbYes Then

' Go To The WhatsApp Search Bar
WScript.Sleep 50
WshShell.SendKeys "{TAB}"


' Go To The Contacts Chat
WScript.Sleep 50
WshShell.SendKeys Contact
WScript.Sleep 50
WshShell.SendKeys "{ENTER}"
WScript.Sleep 200
WScript.Sleep 4000

' The Loop For The Messages
For i = 1 to T
WScript.Sleep 5
WshShell.SendKeys Message
WScript.Sleep 5
WshShell.SendKeys "{ENTER}"
Next

' End Of The Script
WScript.Sleep 3000
MsgBox "MultiMessaging to " + Contact + " Is Done", 1024 + vbSystemModal, "WhatsApp MultiMessage - Dozy.in"

' Canceled Script
Else
MsgBox "Process Has Been Canceled", vbSystemModal, "WhatsApp MultiMessage Canceled"
End If
Else
End If
' Go To DOZY
Set WshShell = WScript.CreateObject("WScript.Shell")
Return = WshShell.Run("http://dozy.in/", 1)

' Loading Time

WScript.Sleep 300

' Go To WhatsApp
' Set WshShell = WScript.CreateObject("WScript.Shell")
' Loading Time