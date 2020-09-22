VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Restart"
      Height          =   495
      Left            =   10680
      TabIndex        =   9
      Top             =   5280
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   100
      TabIndex        =   8
      Top             =   7450
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   10635
      TabIndex        =   7
      Top             =   7080
      Width           =   2200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   5055
      Left            =   10560
      TabIndex        =   3
      Top             =   120
      Width           =   2175
      Begin VB.HScrollBar HScroll4 
         Enabled         =   0   'False
         Height          =   255
         Left            =   100
         TabIndex        =   18
         Top             =   1920
         Width           =   2000
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         LargeChange     =   20
         Left            =   100
         Max             =   100
         TabIndex        =   13
         Top             =   4700
         Value           =   100
         Width           =   2000
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         LargeChange     =   100
         Left            =   105
         Max             =   800
         Min             =   50
         SmallChange     =   10
         TabIndex        =   12
         Top             =   3240
         Value           =   400
         Width           =   2000
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   1000
         Left            =   105
         Max             =   8000
         Min             =   1000
         SmallChange     =   100
         TabIndex        =   11
         Top             =   2640
         Value           =   4000
         Width           =   2000
      End
      Begin VB.CheckBox ChkOrbit 
         Caption         =   "Particale real motion"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox ChkRef 
         Caption         =   "Reflection"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CheckBox ChkAcc 
         Caption         =   "Incident wave"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Surface elevation"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1800
      End
      Begin VB.Label Label6 
         Height          =   375
         Left            =   150
         TabIndex        =   20
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Deep      bottom    Shallow"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   135
         TabIndex        =   19
         Top             =   1680
         Width           =   1905
      End
      Begin VB.Label Label3 
         Caption         =   "Kref"
         Height          =   375
         Left            =   150
         TabIndex        =   16
         Top             =   4500
         Width           =   800
      End
      Begin VB.Label Label2 
         Caption         =   "H"
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "L"
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   2400
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   10635
      TabIndex        =   2
      Top             =   6480
      Width           =   2200
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      Height          =   495
      Left            =   10635
      TabIndex        =   1
      Top             =   5880
      Width           =   2200
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   12240
      Top             =   4920
   End
   Begin VB.PictureBox Picture1 
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7155
      ScaleWidth      =   10140
      TabIndex        =   0
      Top             =   120
      Width           =   10200
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   6360
         Width           =   6855
      End
      Begin VB.Shape Circle1 
         BorderColor     =   &H80000004&
         Height          =   610
         Left            =   3480
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   610
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000002&
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   0
         Left            =   100
         Shape           =   3  'Circle
         Top             =   1250
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000002&
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   0
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1560
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         X1              =   200
         X2              =   9600
         Y1              =   1250
         Y2              =   1250
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         FillColor       =   &H00FFFF00&
         Height          =   60
         Index           =   0
         Left            =   50
         Shape           =   3  'Circle
         Top             =   1250
         Visible         =   0   'False
         Width           =   60
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Timer As Boolean:   Dim MyVar0 As String:   Dim FileO As String:

Dim H As Double:        Dim L As Double:        Dim T As Double

Dim n As Double:        Dim ni As Double:       Dim nr As Double:       Dim p As Single
Dim nrL As Double:      Dim nrT As Double:      Dim nrH As Double:      Dim nrF As Double

Dim NumH As Integer:     Dim IndexH As Single:  Dim Arg As Byte:        Dim dT As Double:
Dim NumV As Integer:     Dim IndexV As Single
Dim NumX As Integer:     Dim IndexX As Single:  Dim NumXa As Integer:   Dim NumXb As Integer:

Dim Vv() As Double:     Dim rV As Double
Dim Xx() As Double:     Dim rX As Double:       Dim a As Double:        Dim b As Double:    Dim Num As Integer
Dim Hh() As Double:

Dim depth As Single:    Dim D As Single:        Dim Gauge1 As Byte:     Dim Xa As Integer
Dim Kref As Single:     Dim Fi As Integer:      Dim Gauge2 As Byte:     Dim Xb As Integer
Dim High As Single:     Dim Low As Single:      Dim Gauge3 As Byte

Dim o As Double         'Radian Frequency
Dim k As Double         'Wave NumXber
Dim c As Double

Dim kv As Single:       Dim kx As Single::      Dim kh As Single:      Dim kvx As Single

Const e = 2.30258:      Const g = 9.81:         Const PI = 3.14159265354:


Private Sub Form_Load()
    
    Form1.Show:             SetTimer hwnd, 1, 1000, AddressOf TimerProc 'auto close msgbox after 2000 ms
    
    Call wave_parameters
    
    IndexH = 48:            ReDim Hh(IndexH) As Double      'Number of horizontal particals (surface)
    IndexV = 12:            ReDim Vv(IndexV) As Double        'Number of vertical particals (orbital partical)
    a = 25:                 b = 12                  'Number of horizontal & vertical particals (orbital partical)
    IndexX = a * b:         ReDim Xx(IndexX) As Double
    
    Call CreateTable_1H(IndexH):            Call CreateTable_1V(IndexV):        Call CreateTable_2X(a, b)
    
    Call opt_check:                         Call ChkOrbit_Click
    
    Shape1(Gauge1).BorderColor = vbRed:     Shape1(Gauge2).BorderColor = vbRed: Shape1(Gauge3).BorderColor = vbRed
    
    MyVar0 = Empty          'variable that is saved to file - surface elevation location Gauge1
    
    'frequency of taking samples is measured in seconds
    Timer1.Interval = 100           'is measured in milisecs - it is set to 10 samples in 1 sec
    dT = Timer1.Interval / 1000     'aquisitioni record is measured in seconds
    
    Timer = True                    'boolean on 1 button set timer1 on/off
    Low = 0:        High = 3000     'duration of animation in dT
    
    With ProgressBar1      'show remaining time of animation
        .Max = High
        .Min = Low
        .Value = High
    End With
    HScroll4.Enabled = False
End Sub

Private Sub CreateTable_1H(IndexH)
    For NumH = 1 To IndexH
            Load Shape1(NumH)
        With Shape1(NumH)                        ' Set locations of surface particles
            .Left = n / 5 + NumH * kh
            .Top = n
            .Visible = True
        End With
        Hh(NumH) = n / 5 + NumH * kh
    Next NumH
End Sub

Private Function CreateTable_1V(IndexZ)
    For NumV = 1 To IndexV
            Load Shape2(NumV)
        With Shape2(NumV)                       'Set locations of vertical particles
            .Left = n / 5 + kv * Gauge3
            .Top = n + H / 2 + NumV * 100
            .Visible = True
        End With
        Vv(NumV) = n / 5 + kv * Gauge3
    Next NumV
End Function

Private Function CreateTable_2X(a, b)
    NumX = 1
    
    For NumXa = 1 To a
        For NumXb = 1 To b:
    
            Load Shape3(NumX)
            
                With Shape3(NumX)             ' Set locations of volume particles
                    .Left = n / 5 + NumXa * kx
                    .Top = n + H / 2 + NumXb * kvx
                    .Visible = True
                End With
                
            Xx(NumX) = n / 5 + NumXa * kx
            
            NumX = NumX + 1
        Next NumXb
    Next NumXa
End Function




Private Sub Timer1_Timer()
    
    For NumH = 1 To IndexH                  'shape1(1) to shape1(IndexH) are moving
        Call Wave(Arg)
    Next NumH
    
    For NumV = 1 To IndexV                  'shape2(1) to shape1(IndexH)IndexV are moving
        Call WaveV(Arg)
    Next NumV
    
    Num = 1
    For Xa = 1 To a
        For Xb = 1 To b                     'create matrix of objects horizontal direction  - a
            Call WaveX(Arg, Xa, Xb, Num)    '                         vertical direction    - b
            Num = Num + 1
        Next Xb
    Next Xa
    
    dT = Round(dT, 1) + Round(Timer1.Interval / 1000, 1)
    
    Label4.Caption = Round(dT, 1) & " " & Shape1(Gauge1).Top & " " & _
                                    Shape1(Gauge2).Top & " " & Shape1(Gauge3).Top
    
    MyVar0 = MyVar0 & vbCrLf & Label4.Caption
    
    If ProgressBar1.Value = 0 Then
        Call cmdSave_Click
    Else
        ProgressBar1.Value = ProgressBar1.Value - 1
    End If
    
End Sub

Private Function Wave(Arg)
    Select Case Arg
        Case 0          'no incident wave
            'if chkref.Value =0
            Shape1(NumH).Top = n:       ChkRef.Enabled = False:      ChkOrbit.Enabled = False
        Case 1          'oscillate around still water level - incident wave
            ChkRef.Enabled = True:      ChkOrbit.Enabled = True
            ni = H / 2 * Sin(o * dT - k * Hh(NumH))
            Shape1(NumH).Top = n + ni
        Case 2          'oscillate around still water level - reflected motion added
            ChkRef.Enabled = True:      ChkOrbit.Enabled = True
            ni = H / 2 * Sin(o * dT - k * Hh(NumH))
            nr = Kref * H / 2 * Sin(o * dT + k * Hh(NumH))
            Shape1(NumH).Top = n + ni + nr
        Case 3
            MsgBox "H  " & Arg
    End Select
End Function

Private Function WaveV(Arg)
    rV = e ^ (-k * (NumV - 1) * kv)
    Select Case Arg
        Case 1      'oscillate around still water level - incident wave
            Shape2(NumV).Left = n / 5 + kv * Gauge3 + rV * H / 2 * Cos(o * dT - k * (Vv(NumV)))
            Shape2(NumV).Top = n + (NumV - 1) * kv + rV * H / 2 * Sin(o * dT - k * Vv(NumV))
            
        Case 2      'oscillate around still water level - reflected motion added
            Shape2(NumV).Left = n / 5 + kv * Gauge3 + rV * H / 2 * Cos(o * dT - k * Vv(NumV)) _
                                + Kref * rV * H / 2 * Cos(o * dT + k * Vv(NumV))
            Shape2(NumV).Top = n + (NumV - 1) * kv + rV * H / 2 * Sin(o * dT - k * Vv(NumV)) _
                                + Kref * rV * H / 2 * Sin(o * dT + k * Vv(NumV))
        Case 0      'no incident wave
            If ChkAcc = 0 Then Call cmdStart_Click   'GoTo lineV
            'Shape2(NumV).Left = n / 5 + kv * Gauge3 + rV * H / 2 * Cos(o * dT - k * Vv(NumV)) _
                                + Kref * rV * H / 2 * Cos(o * dT + k * Vv(NumV))
            'Shape2(NumV).Top = n + (NumV - 1) * kv + rV * H / 2 * Sin(o * dT - k * Vv(NumV)) _
                                + Kref * rV * H / 2 * Sin(o * dT + k * Vv(NumV))
            MsgBox "Err.. V"
    End Select
lineV:
End Function

Private Function WaveX(Arg, Xa, Xb, Num)
    
    rX = e ^ (-k * (Xb - 1) * kx)
    Select Case Arg
        Case 1      'oscillate around still water level - i
            
            Shape3(Num).Left = n / 5 + Xa * kx + rX * H / 2 * Cos(o * dT - k * Xx(Num))
            Shape3(Num).Top = n + (Xb - 1) * kx + rX * H / 2 * Sin(o * dT - k * Xx(Num))
                
        Case 2
            If ChkOrbit = 0 Then GoTo lineX
            Shape3(Num).Left = n / 5 + Xa * kx + rX * H / 2 * Cos(o * dT - k * Xx(Num)) _
                                + Kref * rX * H / 2 * Cos(o * dT + k * Xx(Num))
            'Shape3(Num).Left = n / 5 + Xa * 250 - Kref * rX * H / 2 * Cos(o * dT - k * Xx(Num))
            Shape3(Num).Top = n + (Xb - 1) * kx + rX * H / 2 * Sin(o * dT - k * Xx(Num)) _
                                + Kref * rX * H / 2 * Sin(o * dT + k * Xx(Num))
            'MsgBox Xb
        Case 0
            If ChkOrbit.Value = 1 Or ChkRef.Value = 0 Then GoTo lineX
            MsgBox "X0"
    End Select
lineX:
End Function


Private Sub ChkOrbit_Click()
    If ChkOrbit.Value = 1 Then                'show vertical orbits on 1 line
        For NumV = 1 To IndexV
            Shape2(NumV).Visible = True
        Next NumV
        
        Call Orbit_Circle
        
        For NumX = 1 To IndexX              'show volume movement
            Shape3(NumX).Visible = True
        Next NumX
    Else
        For NumV = 1 To IndexV              'hide all of the above
            Shape2(NumV).Visible = False
        Next NumV
        
        Circle1.Visible = False
        
        For NumX = 1 To IndexX
            Shape3(NumX).Visible = False
        Next NumX
    End If
End Sub
Private Function Orbit_Circle()
        With Circle1                         'add circle to show trajectory
            .Height = H * (1 - rV)
            .Width = H * (1 - rV)
            .Left = n / 5 + kv * Gauge3 - Circle1.Width / 2 + 30
            .Top = n - H / 2 + 20
            .Visible = True
        End With
End Function

Private Function opt_check()
    If ChkAcc.Value = 0 Then
        ChkOrbit.Enabled = 0
    Else
        ChkOrbit.Enabled = 1
    End If
        
    If ChkRef.Value = 0 And ChkAcc.Value = 0 Then               'MsgBox "Dead water"
        Arg = 0
    ElseIf ChkRef.Value = 0 And ChkAcc.Value = 1 Then           'MsgBox "Acc"
        Arg = 1:                    HScroll3.Enabled = False:   Label3.Enabled = False: Call wave_info
    ElseIf ChkRef.Value = 1 And ChkAcc.Value = 1 Then           'MsgBox "Ref+Acc"
        Arg = 2:                    HScroll3.Enabled = True:    Label3.Enabled = True
    Else
        Arg = 3
        'HScroll3.Enabled = False:   Label3.Enabled = False:     Call cmdStart_Click:        Call cmdPause_Click
        MsgBox "No Incident wave - Arg= " & Arg
    End If
End Function




Private Sub ChkAcc_Click()
    Call opt_check
End Sub
Private Sub ChkRef_Click()
    Call opt_check
End Sub





Private Sub cmdStart_Click()
    MyVar0 = Empty
    
    For NumH = 1 To IndexH
            Unload Shape1(NumH)
        Hh(NumH) = 0
    Next NumH
    
    For NumV = 1 To IndexV
            Unload Shape2(NumV)
        Vv(NumV) = 0
    Next NumV
    
    NumX = 1
    For NumXa = 1 To a
        For NumXb = 1 To b
            Unload Shape3(NumX)
            NumX = NumX + 1
        Next NumXb
    Next NumXa
    
    Call Form_Load
End Sub

Private Sub cmdPause_Click()
    If Timer = False Then
        Timer1.Enabled = True
        Timer = True:   cmdPause.Caption = "&Pause":    cmdStart.Enabled = True
    Else
        Timer1.Enabled = False
        Timer = False:   cmdPause.Caption = "&Paused..":    cmdStart.Enabled = False
    End If
End Sub

Private Sub Form_Click()
    Call cmdExit_Click
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub HScroll1_Change()
    Call wave_parameters:      Call cmdStart_Click
End Sub
Private Sub HScroll2_Change()
    Call wave_parameters:      Call Orbit_Circle
End Sub
Private Sub HScroll3_Change()
    Call wave_parameters:
End Sub


Private Sub cmdSave_Click()
    FileO = App.Path & "\" & "Result.txt"
    
    Open FileO For Output As #1
        Print #1, MyVar0
    Close #1
End Sub

Public Function wave_info()
    Form1.Caption = "Linear theory water wave motion - animated by Robert Dolovcak"
    Label1.Caption = "L= " & L / 100 & " m":       Label2.Caption = "H= " & H / 100 & " m"
    Label3.Caption = "Kref= " & Kref
    Label6.Caption = "T= " & Round(T, 2) & " sec" & "    L/H= " & Round(L / H, 1)
End Function

Private Function wave_parameters()
    depth = 5000            'water bed level - distance from uper picture bound
    n = 1250                'still water level
    D = depth - n           'true depth
    
    'wave parameters
    L = HScroll1.Value:     H = HScroll2.Value:          T = (2 * PI * L / 100 / g) ^ 0.5
    o = 2 * PI / T:         k = 2 * PI / L:              c = o / k:                                 Call wave_info
    
    
    
    'position of colored particles - gauge on particle nr#
    Gauge1 = 7:    Gauge2 = 13:    Gauge3 = 35:
    
    Kref = HScroll3.Value / 100     'coefficient of wave reflection
    
    p = 0.99        'procentage of reflected wave that do not change L,T,H,Fi parameters
    Fi = 0          'phase shift of reflected wave pomak ref
    kh = L / 16     'distance between particles of surface plot
    kv = L / 16     'distance between particles of vetical plot
    kx = L / 16     'distance between particles of volume plot
    kvx = L / 16
End Function
