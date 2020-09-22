Attribute VB_Name = "Module1"
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" _
        (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount _
        As Long) As Long

Public Declare Function GetActiveWindow Lib "user32.dll" () As Long

Public Declare Function SetTimer Lib "user32.dll" _
       (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long

Public Declare Function KillTimer Lib "user32.dll" _
        (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long




Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal _
    lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) _
    As Long


Dim strBuff As String * 50
Dim Ahwnd&
Dim Bhwnd&

Public Function LaunchApp(ByVal URL As String) As Long
    On Error Resume Next

    Dim strFile As String
    
    strFile = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, _
        vbNormalFocus)
    
End Function


Sub TimerProc(ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, _
            ByVal lpTimerFunc As Long)
   
   On Error Resume Next
   
   strBuff = String(50, " ")
   
   Ahwnd = GetActiveWindow:    GetClassName Ahwnd, strBuff, 50
   
   If Left(strBuff, 6) = "#32770" Then
       SendKeys "{ENTER}"
   End If
End Sub


