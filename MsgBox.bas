Attribute VB_Name = "Module2"
Option Explicit

Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" _
        (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount _
        As Long) As Long

Declare Function GetActiveWindow Lib "user32.dll" () As Long

Declare Function SetTimer Lib "user32.dll" _
        (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long

Declare Function KillTimer Lib "user32.dll" _
        (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
 

Dim strBuff As String * 50
Dim Ahwnd&


Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, _
            ByVal lpTimerFunc As Long)
   On Error Resume Next
   strBuff = String(50, " ")
   Ahwnd = GetActiveWindow
   GetClassName Ahwnd, strBuff, 50
   If Left(strBuff, 6) = "#32770" Then
       SendKeys "{ENTER}"
   End If
End Sub
