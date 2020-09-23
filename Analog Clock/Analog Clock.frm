VERSION 5.00
Begin VB.Form MainFrm 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5550
   Icon            =   "Analog Clock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6135
      Top             =   5865
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5370
      Top             =   5865
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   18
      Left            =   4425
      MouseIcon       =   "Analog Clock.frx":0742
      MousePointer    =   99  'Custom
      ToolTipText     =   "Help"
      Top             =   5730
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   19
      Left            =   4155
      MouseIcon       =   "Analog Clock.frx":0894
      MousePointer    =   99  'Custom
      ToolTipText     =   "Hide For 10 Seconds"
      Top             =   5910
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   20
      Left            =   4665
      MouseIcon       =   "Analog Clock.frx":09E6
      MousePointer    =   99  'Custom
      ToolTipText     =   "Exit"
      Top             =   5910
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   225
      Index           =   16
      Left            =   2715
      MousePointer    =   15  'Size All
      ToolTipText     =   "Move"
      Top             =   2715
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   14
      Left            =   4155
      MouseIcon       =   "Analog Clock.frx":0B38
      MousePointer    =   99  'Custom
      Top             =   6180
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   15
      Left            =   4665
      MouseIcon       =   "Analog Clock.frx":0C8A
      MousePointer    =   99  'Custom
      Top             =   6180
      Width           =   150
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   2745
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   165
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   2715
      Shape           =   3  'Circle
      Top             =   2715
      Width           =   225
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   151
      X2              =   233
      Y1              =   424
      Y2              =   424
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      BorderWidth     =   4
      Index           =   0
      X1              =   43
      X2              =   125
      Y1              =   424
      Y2              =   424
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   149
      X2              =   234
      Y1              =   409
      Y2              =   409
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      BorderWidth     =   7
      Index           =   0
      X1              =   41
      X2              =   126
      Y1              =   409
      Y2              =   409
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      Index           =   1
      X1              =   151
      X2              =   233
      Y1              =   392
      Y2              =   392
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   9
      Index           =   0
      X1              =   43
      X2              =   125
      Y1              =   392
      Y2              =   392
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   59
      Left            =   6465
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   58
      Left            =   6330
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   57
      Left            =   6165
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   56
      Left            =   6015
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   55
      Left            =   5865
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   54
      Left            =   5715
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   53
      Left            =   5550
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   52
      Left            =   5415
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   51
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   50
      Left            =   5145
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   49
      Left            =   5010
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   48
      Left            =   4860
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   47
      Left            =   4710
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   46
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   45
      Left            =   4410
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   44
      Left            =   4260
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   43
      Left            =   4110
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   42
      Left            =   3945
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   41
      Left            =   3780
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   40
      Left            =   3630
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   39
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   38
      Left            =   3315
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   37
      Left            =   3135
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   36
      Left            =   2925
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   35
      Left            =   2745
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   34
      Left            =   2580
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   33
      Left            =   2415
      Shape           =   3  'Circle
      Top             =   5175
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   32
      Left            =   2220
      Shape           =   3  'Circle
      Top             =   5175
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   31
      Left            =   2025
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   30
      Left            =   1845
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   29
      Left            =   1665
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   28
      Left            =   1485
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   27
      Left            =   1335
      Shape           =   3  'Circle
      Top             =   5250
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   26
      Left            =   1170
      Shape           =   3  'Circle
      Top             =   5250
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   25
      Left            =   975
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   24
      Left            =   810
      Shape           =   3  'Circle
      Top             =   5265
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   23
      Left            =   660
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   22
      Left            =   6450
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   21
      Left            =   6255
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   20
      Left            =   6060
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   19
      Left            =   5895
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   18
      Left            =   5730
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   17
      Left            =   5340
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   16
      Left            =   5190
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   15
      Left            =   5055
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   14
      Left            =   4875
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   13
      Left            =   4665
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   12
      Left            =   4425
      Shape           =   3  'Circle
      Top             =   5460
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   11
      Left            =   4275
      Shape           =   3  'Circle
      Top             =   5460
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   10
      Left            =   4095
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   9
      Left            =   3930
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   8
      Left            =   3750
      Shape           =   3  'Circle
      Top             =   5460
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   7
      Left            =   3540
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   6
      Left            =   3255
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   5
      Left            =   3015
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   4
      Left            =   2835
      Shape           =   3  'Circle
      Top             =   5505
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   3
      Left            =   2610
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   2
      Left            =   2325
      Shape           =   3  'Circle
      Top             =   5505
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   1
      Left            =   2100
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   0
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   150
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PI = 3.141592654

Dim Lastx As Integer
Dim Lasty As Integer
Dim Saveleft As Variant
Dim Savetop As Variant
Dim DelayCounter As Long
Dim TopMost As Variant
Dim Startup As Variant
Public ShowHelp As Variant

Const LWA_COLORKEY = &H1
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const BM_SETSTATE = &HF3

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Sub Form_Load()
  Dim Ret As Long
  Dim CLR As Long
  Me.Hide
  For i = 0 To 59
    If i Mod 5 = 0 Then
      Shape4(i).Left = 188 + Cos(i * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
      Shape4(i).Top = 188 + Sin(i * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
    Else
      Shape4(i).Left = 188 + Cos(i * 2 * PI / 60 - (0.5 * PI)) * 100 - 2.5
      Shape4(i).Top = 188 + Sin(i * 2 * PI / 60 - (0.5 * PI)) * 100 - 2.5
    End If
    Shape4(i).BorderColor = vbBlack
  Next i
  Image2(18).Left = 188 + Cos(0 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(18).Top = 188 + Sin(0 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(19).Left = 188 + Cos(50 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(19).Top = 188 + Sin(50 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(20).Left = 188 + Cos(10 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(20).Top = 188 + Sin(10 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(14).Left = 188 + Cos(40 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(14).Top = 188 + Sin(40 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(15).Left = 188 + Cos(20 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(15).Top = 188 + Sin(20 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  
  CLR = RGB(0, 0, 255) 'this color is the color that will be transparent
  'Set the window style to 'Layered'
  Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
  Ret = Ret Or WS_EX_LAYERED
  SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
  'Set the opacity of the layered window to 128
  SetLayeredWindowAttributes Me.hWnd, CLR, 0, LWA_COLORKEY
  OpenFile
  SetTopMost
  SetStartUp
  Me.Show
End Sub
Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  Select Case Index
    Case 14
      Startup = Not Startup
      SetStartUp
    Case 15
      TopMost = Not TopMost
      SetTopMost
      SaveFile
    Case 16
      Lastx = x
      Lasty = y
    Case 18
      HelpFrm.Show
    Case 19
      MainFrm.Hide
      DelayCounter = 0
      Timer2.Interval = 100
    Case 20
      End
  End Select
End Sub
Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  Select Case Index
    Case 16
      If Button = 1 Then
        MainFrm.Left = MainFrm.Left + (x - Lastx)
        MainFrm.Top = MainFrm.Top + (y - Lasty)
      End If
  End Select
End Sub
Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  SaveFile
End Sub
Private Sub SetStartUp()
  If (Startup) Then
    Image2(14).ToolTipText = "Run Manually"
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "Transparent Analog Clock", App.Path & "\" & App.EXEName & ".exe"
    SaveFile
  Else
    Image2(14).ToolTipText = "Run When Windows Starts Up"
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "Transparent Analog Clock", "<NonRun>"
    SaveFile
  End If
End Sub
Private Sub SetTopMost()
  If TopMost Then
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Image2(15).ToolTipText = "Make Not Always On Top"
  Else
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Image2(15).ToolTipText = "Make Always On Top"
  End If
End Sub
Private Sub OpenFile()
  On Error GoTo There_Is_No_File
  Open "Analog Clock.dat" For Input Access Read As #1
    Line Input #1, Saveleft
    Line Input #1, Savetop
    Line Input #1, TopMost
    Line Input #1, Startup
    Line Input #1, ShowHelp
  Close #1
  GoTo The_File_Exist
There_Is_No_File:
  Saveleft = 0
  Savetop = 0
  TopMost = True
  Startup = False
  ShowHelp = 1
  SaveFile
The_File_Exist:
  MainFrm.Left = Saveleft
  MainFrm.Top = Savetop
  HelpFrm.Check1.Value = ShowHelp
  If ShowHelp Then HelpFrm.Show
End Sub
Public Sub SaveFile()
  Open "Analog Clock.dat" For Output As #1
    Print #1, MainFrm.Left
    Print #1, MainFrm.Top
    Print #1, TopMost
    Print #1, Startup
    Print #1, ShowHelp
  Close #1
End Sub
Private Sub Timer1_Timer()
  Dim Tim As Long
  Tim = Int(Timer)
  For i = 0 To 1
    'Hour
    Line1(i).X1 = 188
    Line1(i).Y1 = 188
    Line1(i).X2 = 188 + Cos((Tim Mod 43200) * 2 * PI / 43200 - (0.5 * PI)) * 60
    Line1(i).Y2 = 188 + Sin((Tim Mod 43200) * 2 * PI / 43200 - (0.5 * PI)) * 60
    'Minute
    Line2(i).X1 = 188
    Line2(i).Y1 = 188
    Line2(i).X2 = 188 + Cos((Tim \ 60 Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 90
    Line2(i).Y2 = 188 + Sin((Tim \ 60 Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 90
    'Second
    Line3(i).X1 = 188 - Cos((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 15
    Line3(i).Y1 = 188 - Sin((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 15
    Line3(i).X2 = 188 + Cos((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 90
    Line3(i).Y2 = 188 + Sin((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 90
  Next i
End Sub
Private Sub Timer2_Timer()
  DelayCounter = DelayCounter + 1
  If DelayCounter >= 100 Then
     MainFrm.Show
     Timer2.Interval = 0
  End If
End Sub
