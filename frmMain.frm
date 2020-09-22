VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CleanSweepPro"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAbout 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   600
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Label lblAboutArray 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0BC2
         Height          =   855
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   5535
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   5160
         Picture         =   "frmMain.frx":0CFA
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblAboutArray 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":18BC
         Height          =   615
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label lblAboutArray 
         BackStyle       =   0  'Transparent
         Caption         =   "Chase Gale"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblAboutArray 
         BackStyle       =   0  'Transparent
         Caption         =   "Authored By:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblAboutArray 
         BackStyle       =   0  'Transparent
         Caption         =   "Special Thanks to:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblAboutArray 
         BackStyle       =   0  'Transparent
         Caption         =   "Christoph von Wittich"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   8
         Top             =   460
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   6375
      Begin VB.ListBox List1 
         BackColor       =   &H00C0FFFF&
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   6135
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ack! Delete all of it!"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblWinTemp 
         BackStyle       =   0  'Transparent
         Caption         =   "Click here to locate your Windows' Temp Folder!"
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmMain.frx":196F
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1920
         Width           =   3615
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   5760
      MouseIcon       =   "frmMain.frx":1AC1
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   30
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Internet Cache Listing:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   6480
      MouseIcon       =   "frmMain.frx":1C13
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      FillColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   6840
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declares for Form Dragging
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
  lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'Constants for form dragging
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'Create new object from the clsSideBar class (to make the neato sidebar)
Dim sidebar As New clsSideBar


Private Sub cmdDelete_Click()
'Call the DeleteCache Sub in basWinInet(.bas)
DeleteCache
End Sub


Private Sub Form_Load()
'Call the EnumerateCache sub in basWinInet that populates the listBox
EnumerateCache
'Show the form after the listbox has been populated, this is done because of
'Required form elements used in the sidebar class.
Show
'Call the sidebar class, clear any exsisting bar (just good practice) and create
'the new one.
sidebar.RemoveTitleBar
sidebar.Create "Clean Sweep Pro", 1, 1, QBColor(15), QBColor(8), "Tahoma", 16
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call FormDrag(Me)
End Sub

Private Sub FormDrag(Frm As Form)
    ReleaseCapture
    Call SendMessage(Frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call FormDrag(frmMain)
End Sub

Private Sub FrameAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call FormDrag(frmMain)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call FormDrag(frmMain)
End Sub

Private Sub Label2_Click()
Unload Me 'Duh?
End Sub

Private Sub Label3_Click()
'Small if then to hide/show stuff based on caption of label
If Label3.Caption = "About..." Then
    Frame1.Visible = False
    FrameAbout.Visible = True
    Label3.Caption = "Back..."
    Label1.Caption = "About " & App.Title
Else
    Frame1.Visible = True
    FrameAbout.Visible = False
    Label3.Caption = "About..."
    Label1.Caption = "Current Internet Cache Listing:"
End If
End Sub

Private Sub lblAboutArray_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call FormDrag(frmMain)
End Sub

Private Sub lblWinTemp_Click()
'Make a msgbox that gets data from the 'Gettmppath' function in basWindowsTemp
Call MsgBox("Windows Temp. Path: " & GetTmpPath, vbInformation)
End Sub
