Attribute VB_Name = "Module1"
Declare Function GetTickCount Lib "kernel32" () As Long


Sub timeout(length)
starttime = Timer
Do While Timer - starttime < length
DoEvents
Loop
End Sub

