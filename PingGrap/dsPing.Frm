VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pinger log + graphic"
   ClientHeight    =   8310
   ClientLeft      =   1740
   ClientTop       =   2025
   ClientWidth     =   9270
   Icon            =   "dsPing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8310
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Graphic :"
      Height          =   4575
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   9015
      Begin VB.HScrollBar Scroll 
         Height          =   255
         LargeChange     =   100
         Left            =   720
         Max             =   0
         SmallChange     =   10
         TabIndex        =   11
         Top             =   4200
         Width           =   8055
      End
      Begin VB.PictureBox GrapHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   3750
         Left            =   720
         ScaleHeight     =   9999
         ScaleMode       =   0  'User
         ScaleWidth      =   537
         TabIndex        =   10
         Top             =   330
         Width           =   8055
         Begin VB.PictureBox Graphic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            ForeColor       =   &H000000FF&
            Height          =   3750
            Left            =   0
            ScaleHeight     =   2000
            ScaleMode       =   0  'User
            ScaleWidth      =   537
            TabIndex        =   15
            Top             =   0
            Width           =   8055
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   19
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label2 
         Caption         =   "P I N G - M O N I T O R"
         Height          =   2355
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   135
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "0         -"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   3945
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2000"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Option :"
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   9015
      Begin VB.CheckBox Check1 
         Caption         =   "Dont show 'timed out' pings in the graphic"
         Height          =   195
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Value           =   1  'Checked
         Width           =   5655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   16
         Text            =   "100"
         Top             =   1560
         Width           =   4575
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   5640
         Top             =   600
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "1000"
         Top             =   1080
         Width           =   4575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "S&top"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6960
         TabIndex        =   7
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Text            =   "212.120.66.200"
         Top             =   600
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Start"
         Default         =   -1  'True
         Height          =   375
         Left            =   6960
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Max ping :"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Interval :"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IP :"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Log :"
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton Command2 
         Caption         =   "&Clear log + graphic"
         Height          =   375
         Left            =   6960
         TabIndex        =   2
         Top             =   900
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Maximum ping returned :"
         Height          =   195
         Index           =   0
         Left            =   6840
         TabIndex        =   24
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Minimum ping returned :"
         Height          =   195
         Index           =   1
         Left            =   6885
         TabIndex        =   23
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   8640
         TabIndex        =   22
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   8640
         TabIndex        =   21
         Top             =   600
         Width           =   90
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X1 As Long, Y1 As Long

Private Sub Command1_Click()
If ValuesFaild = True Then
Command2_Click
Command1.Enabled = False
Command3.Enabled = True
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Graphic.ScaleHeight = Val(Text3)
SetIp Text1
Label1(0).Caption = Text3
Timer1.Interval = Text2
Timer1.Enabled = True
End If
End Sub

Private Sub Command2_Click()
Label4(0).Caption = 0
Label4(1).Caption = Text3
List1.Clear
Graphic.Cls
Graphic.Width = GrapHolder.ScaleWidth
Scroll.Value = 0
Scroll.Max = 0
X1 = 0
Y1 = Graphic.ScaleHeight
End Sub

Private Sub Command3_Click()
Command1.Enabled = True
Command3.Enabled = False
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Form_Load()
Command2_Click
End Sub

Private Sub Scroll_Change()
Graphic.Left = -Scroll.Value
End Sub

Private Sub Timer1_Timer()
Dim Ret As PingReturn
Dim X2 As Long, Y2 As Long
Ret = Pinger(20, Val(Text3))
List1.AddItem Ret.Discriptionn
List1.ListIndex = List1.ListCount - 1
X2 = X1 + 2
Y2 = Graphic.ScaleHeight - Ret.Ping
If Ret.Errors = True And Check1.Value = 1 Then Exit Sub
If Label4(0).Caption < Ret.Ping Then Label4(0).Caption = Ret.Ping
If Label4(1).Caption > Ret.Ping Then Label4(1).Caption = Ret.Ping
Graphic.Line (X1, Y1)-(X2, Y2), vbRed
X1 = X2
Y1 = Y2
CalcFieldandScroll X1 + 4
End Sub

Function ValuesFaild() As Boolean
If Len(Text1) = 0 Then MsgBox "Can't start with out an ip adress!", vbInformation + vbOKOnly, "Error": Text1.SetFocus: Exit Function
If Len(Text2) = 0 Then MsgBox "Can't start with out an interval (Defuald : 1000)", vbInformation + vbOKOnly, "Error": Text2.SetFocus: Exit Function
If IsNumeric(Text2) = False Then MsgBox "Can't start with out an falid interval (Defuald : 1000)", vbInformation + vbOKOnly, "Error": Text2.SetFocus: Exit Function
If Len(Text3) = 0 Then MsgBox "Can't start with out an max ping value (Defuald : 1000)", vbInformation + vbOKOnly, "Error": Text2.SetFocus: Exit Function
If IsNumeric(Text3) = False Then MsgBox "Can't start with out an falid max ping value (Defuald : 1000)", vbInformation + vbOKOnly, "Error": Text2.SetFocus: Exit Function
ValuesFaild = True
End Function

Function CalcFieldandScroll(WidthNeed As Long)
Dim Ret As Boolean
If Scroll.Value + 20 > Scroll.Max Then Ret = True
If WidthNeed > Graphic.ScaleWidth Then
Graphic.Width = WidthNeed
Scroll.Max = Graphic.Width - GrapHolder.ScaleWidth
DoEvents
Scroll.Refresh
If Ret = True Then Scroll.Value = Scroll.Max
End If
End Function
