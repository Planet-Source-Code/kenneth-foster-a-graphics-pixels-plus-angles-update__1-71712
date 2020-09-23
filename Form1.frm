VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2640
   ClientLeft      =   2790
   ClientTop       =   2100
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   258
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      TabIndex        =   3
      Top             =   -30
      Width           =   3810
      Begin VB.CheckBox chkP_T 
         Caption         =   "Twips"
         Height          =   195
         Left            =   1470
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   240
         Left            =   1470
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   300
         Width           =   690
      End
      Begin VB.HScrollBar hsTrans 
         Height          =   210
         LargeChange     =   10
         Left            =   15
         Max             =   225
         Min             =   50
         SmallChange     =   5
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   555
         Value           =   140
         Width           =   3780
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "EXIT"
         Height          =   240
         Left            =   15
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   300
         Width           =   675
      End
      Begin VB.CheckBox chkAngle 
         Caption         =   "Show Point 2"
         Height          =   300
         Left            =   30
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   15
         Width           =   1245
      End
      Begin VB.Label pt2Angle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "180"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2220
         TabIndex        =   10
         Top             =   285
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label pt1Angle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3030
         TabIndex        =   9
         Top             =   285
         Width           =   750
      End
      Begin VB.Label pt2Leng 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2220
         TabIndex        =   8
         Top             =   30
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label pt1Leng 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3030
         TabIndex        =   7
         Top             =   30
         Width           =   750
      End
   End
   Begin VB.PictureBox pt2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   1080
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   2
      Top             =   1740
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox ptMid 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   1695
      Picture         =   "Form1.frx":0043
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   0
      Top             =   1740
      Width           =   105
   End
   Begin VB.PictureBox pt1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   2265
      Picture         =   "Form1.frx":0086
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   1
      Top             =   1740
      Width           =   105
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   113
      X2              =   75
      Y1              =   119
      Y2              =   119
   End
   Begin VB.Line Line1 
      X1              =   119
      X2              =   154
      Y1              =   119
      Y2              =   119
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Kenneth Foster
'Use or abuse as you like
'Jan 2009

Option Explicit
Private Type Pos
   X As Integer
   Y As Integer
End Type

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&
'set transparancy of form
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'form drag
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'stay on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const conHwndTopmost = -1
Private Const conSwpNoActivate = &H10
Private Const conSwpShowWindow = &H40

' used to find angles
Dim Com As Pos
Dim Aa As Pos
Dim Bb As Pos

'drag and drop anywhere to move form
'select a point and use arrow keys to nudge it
'use tab button to advance thru the points
'I left the draw grid code here, but disabled it because it flickers to bad
'maybe someone can fix that

Private Sub Form_Activate()
   Dim NormalWindowStyle As Long
   Dim HWD As Long
   NormalWindowStyle = GetWindowLong(HWD, GWL_EXSTYLE)
   SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
   SetLayeredWindowAttributes Me.hwnd, 0, (150), LWA_ALPHA
   SetWindowPos hwnd, conHwndTopmost, Screen.Width / 80, Screen.Height / 75, 260, 181, conSwpNoActivate Or conSwpShowWindow
End Sub

Private Sub Form_Load()

   pt1Leng.Caption = Round(Abs(Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)))
   
   'set some properties for point boxes
   pt1.ScaleWidth = 5
   pt1.ScaleHeight = 5
   pt2.ScaleWidth = 5
   pt2.ScaleHeight = 5
   ptMid.ScaleWidth = 5
   ptMid.ScaleHeight = 5
   cmdReset_Click      'make sure the points are positioned correctly
End Sub

Private Sub Form_DblClick()
   Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   FormDrag Me
End Sub

Private Sub FormDrag(TheForm As Object)
   On Local Error Resume Next
   ReleaseCapture
   SendMessage TheForm.hwnd, &HA1, 2, 0&
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub cmdReset_Click()
   'reset positions of points and lines to home
   Line2.X1 = 120
   Line2.X2 = 75
   Line2.Y1 = 120
   Line2.Y2 = 120
   
   pt1.Top = 116
   pt1.Left = 151
   ptMid.Top = 116
   ptMid.Left = 113
   pt2.Top = 116
   pt2.Left = 72
   
   Line1.X1 = ptMid.Left + ptMid.Width / 2
   Line1.Y1 = ptMid.Top + ptMid.Height / 2
   Line1.X2 = pt1.Left + pt1.Width / 2
   Line1.Y2 = pt1.Top + pt1.Height / 2
   
   If chkP_T.Value = 0 Then
      pt2Leng.Caption = Round(Abs(Sqr((Line2.X1 - Line2.X2) ^ 2 + (Line2.Y1 - Line2.Y2) ^ 2)))
   Else
      pt2Leng.Caption = Round(Abs(Sqr((Line2.X1 - Line2.X2) ^ 2 + (Line2.Y1 - Line2.Y2) ^ 2) * 15))
   End If
   pt2Angle.Caption = "180"
   
   If chkP_T.Value = 0 Then
      pt1Leng.Caption = Round(Abs(Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)))
   Else
      pt1Leng.Caption = Round(Abs(Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2) * 15))
   End If
   pt1Angle.Caption = "0"
   Me.Width = 3900
   Me.Height = 2720
End Sub

Private Sub cmdExit_Click()
   Unload Me
   End
End Sub

Private Sub chkAngle_Click()
   'make visible
   pt2.Visible = Not pt2.Visible
   Line2.Visible = pt2.Visible
   pt2Leng.Visible = pt2.Visible
   pt2Angle.Visible = pt2.Visible
   'make source of both lines the same point
   'Line2.X1 = Line1.X1
   'Line2.Y1 = Line1.Y1
   
   'value points to find angle
   Aa.X = Line1.X2
   Aa.Y = Line1.Y2
   Com.X = Line1.X1
   Com.Y = Line1.Y1
   pt1Angle.Caption = GetAngle(Aa, Com)
   
   'calculate length of the line
   If chkP_T.Value = 0 Then
      pt2Leng.Caption = Round(Abs(Sqr((Line2.X1 - Line2.X2) ^ 2 + (Line2.Y1 - Line2.Y2) ^ 2)))
   Else
      pt2Leng.Caption = Round(Abs(Sqr((Line2.X1 - Line2.X2) ^ 2 + (Line2.Y1 - Line2.Y2) ^ 2) * 15))
   End If
  
End Sub

Private Sub chkP_T_Click()
   If chkP_T.Value = 0 Then
         pt1Leng.Caption = Round(Abs(Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)))
         pt2Leng.Caption = Round(Abs(Sqr((Line2.X1 - Line2.X2) ^ 2 + (Line2.Y1 - Line2.Y2) ^ 2)))
      Else
         pt1Leng.Caption = Round(Abs(Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2) * 15))
         pt2Leng.Caption = Round(Abs(Sqr((Line2.X1 - Line2.X2) ^ 2 + (Line2.Y1 - Line2.Y2) ^ 2) * 15))
      End If
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   FormDrag Me
End Sub

Private Function GetAngle(Pos As Pos, CenterPos As Pos) As Integer   'I borrowed this function from someone.
   'Returns the angle between two points in
   '     degrees
   Dim intA As Integer
   Dim intB As Integer
   Dim intC As Integer
   Dim PI As Double
   
   PI = Atn(1) * 4
   intB = Abs(CenterPos.X - Pos.X) 'distance is always positive-->abs()
   intC = Abs(CenterPos.Y - Pos.Y)
   
   If intB <> 0 Then 'don't divide by zero ...
   GetAngle = Atn(intC / intB) * 180 / PI
End If

If Pos.X < CenterPos.X Then
   'the point is at the left of CenterPos
   If Pos.Y = CenterPos.Y Then GetAngle = 180
   
   If Pos.Y < CenterPos.Y Then
      GetAngle = 180 - GetAngle
   End If
   
   If Pos.Y > CenterPos.Y Then
      GetAngle = 180 + GetAngle
   End If
End If

If Pos.X > CenterPos.X Then
   'the point is at the right of CenterPos
   If Pos.Y > CenterPos.Y Then
      GetAngle = 360 - GetAngle
   End If
End If

If Pos.X = CenterPos.X Then
   
   If Pos.Y < CenterPos.Y Then
      GetAngle = 90
   End If
   
   If Pos.Y > CenterPos.Y Then
      GetAngle = 270
   End If
End If
'be sure the GetAngle is between [0,360]
GetAngle = Abs(GetAngle Mod 360)
End Function

Private Sub hsTrans_Change()
   SetLayeredWindowAttributes Me.hwnd, 0, (hsTrans.Value), LWA_ALPHA
End Sub

Private Sub hsTrans_Scroll()
   hsTrans_Change
End Sub

Private Sub ptMid_GotFocus()
   ptMid.BackColor = &HFFFF&
End Sub

Private Sub ptMid_LostFocus()
   ptMid.BackColor = vbWhite
End Sub

Private Sub ptMid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      ptMid.Left = ptMid.Left + X
      ptMid.Top = ptMid.Top + Y
      Line1.X1 = ptMid.Left + ptMid.Width / 2
      Line1.Y1 = ptMid.Top + ptMid.Height / 2
      Line1.X2 = pt1.Left + pt1.Width / 2
      Line1.Y2 = pt1.Top + pt1.Height / 2
      Line2.X1 = ptMid.Left + ptMid.Width / 2
      Line2.Y1 = ptMid.Top + ptMid.Height / 2
      If pt2.Visible = True Then
         'make source of both lines the same point
         Bb.X = Line2.X2
         Bb.Y = Line2.Y2
         Com.X = Line1.X1
         Com.Y = Line1.Y1
         pt2Angle.Caption = GetAngle(Bb, Com)
      End If
      'calculate length of line using trig formula for right triangle and format it
      If chkP_T.Value = 0 Then
         pt1Leng.Caption = Round(Abs(Sqr(Abs(Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)))
         pt2Leng.Caption = Round(Abs(Sqr((Line2.X1 - Line2.X2) ^ 2 + (Line2.Y1 - Line2.Y2) ^ 2)))
      Else
         pt1Leng.Caption = Round(Abs(Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2) * 15))
         pt2Leng.Caption = Round(Abs(Sqr((Line2.X1 - Line2.X2) ^ 2 + (Line2.Y1 - Line2.Y2) ^ 2) * 15))
      End If
      'get values to figure angle
      Aa.X = Line1.X2
      Aa.Y = Line1.Y2
      Com.X = Line1.X1
      Com.Y = Line1.Y1
      pt1Angle.Caption = GetAngle(Aa, Com)
   End If
End Sub

Private Sub ptMid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyRight Then
      ptMid.Left = ptMid.Left + 1
      Line1.X1 = ptMid.Left + ptMid.Width / 2
   End If
   If KeyCode = vbKeyLeft Then
      ptMid.Left = ptMid.Left - 1
      Line1.X1 = ptMid.Left + ptMid.Width / 2
   End If
   If KeyCode = vbKeyUp Then
      ptMid.Top = ptMid.Top - 1
      Line1.Y1 = ptMid.Top + ptMid.Height / 2
   End If
   If KeyCode = vbKeyDown Then
      ptMid.Top = ptMid.Top + 1
      Line1.Y1 = ptMid.Top + ptMid.Height / 2
   End If
    If pt2.Visible = True Then
         'make source of both lines the same point
         Line2.X1 = Line1.X1
         Line2.Y1 = Line1.Y1
         Bb.X = Line2.X2
         Bb.Y = Line2.Y2
         pt2Angle.Caption = GetAngle(Bb, Com)
      End If
   ptMid_MouseMove 1, 0, 0, 0
   If Shift = 1 Then
      pt2_KeyDown KeyCode, 0
      pt1_KeyDown KeyCode, 0
   End If
End Sub

Private Sub pt1_GotFocus()
   pt1.BackColor = &H8080FF
End Sub

Private Sub pt1_LostFocus()
   pt1.BackColor = vbWhite
End Sub

Private Sub pt1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      pt1.Left = pt1.Left + X
      pt1.Top = pt1.Top + Y
      Line1.X2 = pt1.Left + pt1.Width / 2
      Line1.Y2 = pt1.Top + pt1.Height / 2
       'calculate length of line using trig formula for right triangle and format it
      If chkP_T.Value = 0 Then
         pt1Leng.Caption = Round(Abs(Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)))
      Else
         pt1Leng.Caption = Round(Abs(Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2) * 15))
      End If
      'get values to figure angle
      Aa.X = Line1.X2
      Aa.Y = Line1.Y2
      Com.X = Line1.X1
      Com.Y = Line1.Y1
      pt1Angle.Caption = GetAngle(Aa, Com)
   End If
End Sub

Private Sub pt1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyRight Then
      pt1.Left = pt1.Left + 1
      Line1.X2 = pt1.Left + pt1.Width / 2
   End If
   If KeyCode = vbKeyLeft Then
      pt1.Left = pt1.Left - 1
      Line1.X2 = pt1.Left + pt1.Width / 2
   End If
   If KeyCode = vbKeyUp Then
      pt1.Top = pt1.Top - 1
      Line1.Y2 = pt1.Top + pt1.Height / 2
   End If
   If KeyCode = vbKeyDown Then
      pt1.Top = pt1.Top + 1
      Line1.Y2 = pt1.Top + pt1.Height / 2
   End If
   pt1_MouseMove 1, 0, 0, 0
   If Shift = 1 Then
      pt2_KeyDown KeyCode, 0
      ptMid_KeyDown KeyCode, 0
   End If
End Sub

Private Sub pt2_GotFocus()
   pt2.BackColor = &HFF8080
End Sub

Private Sub pt2_LostFocus()
   pt2.BackColor = vbWhite
End Sub

Private Sub pt2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      pt2.Left = pt2.Left + X
      pt2.Top = pt2.Top + Y
      Line2.X2 = pt2.Left + pt2.Width / 2
      Line2.Y2 = pt2.Top + pt2.Height / 2
      Line2.X1 = ptMid.Left + ptMid.Width / 2
      Line2.Y1 = ptMid.Top + ptMid.Height / 2
       'calculate length of line using trig formula for right triangle and format it
      If chkP_T.Value = 0 Then
         pt2Leng.Caption = Round(Abs(Sqr((Line2.X1 - Line2.X2) ^ 2 + (Line2.Y1 - Line2.Y2) ^ 2)))
      Else
         pt2Leng.Caption = Round(Abs(Sqr((Line2.X1 - Line2.X2) ^ 2 + (Line2.Y1 - Line2.Y2) ^ 2) * 15))
      End If
      'get values to figure angle
      Bb.X = Line2.X2
      Bb.Y = Line2.Y2
      Com.X = Line2.X1
      Com.Y = Line2.Y1
      pt2Angle.Caption = GetAngle(Bb, Com)
   End If
End Sub

Private Sub pt2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyRight Then
      pt2.Left = pt2.Left + 1
      Line2.X2 = pt2.Left + pt2.Width / 2
   End If
   If KeyCode = vbKeyLeft Then
      pt2.Left = pt2.Left - 1
      Line2.X2 = pt2.Left + pt2.Width / 2
   End If
   If KeyCode = vbKeyUp Then
      pt2.Top = pt2.Top - 1
      Line2.Y2 = pt2.Top + pt2.Height / 2
   End If
   If KeyCode = vbKeyDown Then
      pt2.Top = pt2.Top + 1
      Line2.Y2 = pt2.Top + pt2.Height / 2
   End If
   pt2_MouseMove 1, 0, 0, 0
   If Shift = 1 Then
      pt1_KeyDown KeyCode, 0
      ptMid_KeyDown KeyCode, 0
   End If
End Sub

