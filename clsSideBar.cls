VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSideBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum sbAlign
 sbAlignLeft = 1
 sbAlignRight = 2
End Enum

Public Enum sbFillStyle
 sbFilled = 1
 sbSmooth = 2
End Enum

Private sFillStyle As sbFillStyle
Private sbFontName As String
Private sbText As String
Private sbForeColor As Long
Private sbBackColor As Long
Private sbFontSize As Integer
Private sbActive As Boolean
Private sAlign As sbAlign
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const sBarWidth = 20

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Dim oldMode%
Dim BckupCol&

Public Property Get FillStyle() As sbFillStyle
 FillStyle = sFillStyle
End Property
Public Property Let FillStyle(sBarFillStyle As sbFillStyle)
 sFillStyle = sBarFillStyle
 Refresh
End Property
Public Property Get FontName() As String
 FontName = sbFontName
End Property
Public Property Let FontName(sBarFontName As String)
 sbFontName = sBarFontName
 Refresh
End Property
Public Property Get Text() As String
 Text = sbText
End Property
Public Property Let Text(sBarText As String)
 sbText = sBarText
 Refresh
End Property
Public Property Get BackColor() As Long
 BackColor = sbBackColor
End Property
Public Property Let BackColor(sBarBackColor As Long)
 sbBackColor = sBarBackColor
 Refresh
End Property
Public Property Get ForeColor() As Long
 ForeColor = sbForeColor
End Property
Public Property Let ForeColor(sBarForeColor As Long)
 sbForeColor = sBarForeColor
 Refresh
End Property
Public Property Get FontSize() As Integer
 FontSize = sbFontSize
End Property
Public Property Let FontSize(sBarFontSize As Integer)
 sbFontSize = sBarFontSize
 Refresh
End Property
Public Property Get Align() As sbAlign
 Align = sAlign
End Property
Public Property Let Align(sBarAlign As sbAlign)
 sAlign = sBarAlign
 Refresh
End Property
Public Sub Create(sBarText As String, sBarFillStyle As sbFillStyle, sBarAlign As sbAlign, sBarForeColor As Long, sBarBackColor As Long, sBarFontName As String, sBarFontSize As Integer)
 sbText = sBarText: sFillStyle = sBarFillStyle
 sAlign = sBarAlign: sbForeColor = sBarForeColor
 sbBackColor = sBarBackColor: sbFontSize = sBarFontSize
 sbFontName = sBarFontName
 sbActive = True
 Refresh
End Sub
Public Sub Destroy()
 oldMode% = Screen.ActiveForm.ScaleMode
 Screen.ActiveForm.ScaleMode = 3
 Screen.ActiveForm.Line (0, 0)-(sBarWidth, Screen.ActiveForm.Height), Screen.ActiveForm.BackColor, BF
 Screen.ActiveForm.Line (Screen.ActiveForm.ScaleWidth - sBarWidth + 1, 0)-(Screen.ActiveForm.ScaleWidth, Screen.ActiveForm.Height), Screen.ActiveForm.BackColor, BF
 Screen.ActiveForm.ScaleMode = oldMode%
 sbActive = False
End Sub
Private Sub Farbverlauf(col1&)
 Const intBLUESTART% = 255
 Const intBLUEEND% = 0
 Const intBANDHEIGHT% = 2
 Dim sngBlueCur As Single
 Dim sngBlueStep As Single
 Dim intFormHeight As Integer
 Dim intFormWidth As Integer
 Dim intY As Integer
 intFormHeight = Screen.ActiveForm.ScaleHeight
 If intFormHeight < 1 Then Exit Sub
 intFormWidth = sBarWidth
 sngBlueStep = intBANDHEIGHT * (intBLUEEND - intBLUESTART) / intFormHeight
 sngBlueCur = intBLUESTART
 For intY = 0 To intFormHeight Step intBANDHEIGHT
  Select Case col1&
   Case QBColor(1)
    If sAlign = 1 Then
     Screen.ActiveForm.Line (-1, intY - 1)-(intFormWidth, intY + intBANDHEIGHT), RGB(0, 0, sngBlueCur), BF
    Else
     Screen.ActiveForm.Line (Screen.ActiveForm.ScaleWidth - sBarWidth + 1, intY - 1)-(Screen.ActiveForm.ScaleWidth, intY + intBANDHEIGHT), RGB(0, 0, sngBlueCur), BF
    End If
   Case QBColor(2)
    If sAlign = 1 Then
     Screen.ActiveForm.Line (-1, intY - 1)-(intFormWidth, intY + intBANDHEIGHT), RGB(0, sngBlueCur, 0), BF
    Else
     Screen.ActiveForm.Line (Screen.ActiveForm.ScaleWidth - sBarWidth + 1, intY - 1)-(Screen.ActiveForm.ScaleWidth, intY + intBANDHEIGHT), RGB(0, sngBlueCur, 0), BF
    End If
   Case QBColor(4)
    If sAlign = 1 Then
    Screen.ActiveForm.Line (-1, intY - 1)-(intFormWidth, intY + intBANDHEIGHT), RGB(sngBlueCur, 0, 0), BF
    Else
     Screen.ActiveForm.Line (Screen.ActiveForm.ScaleWidth - sBarWidth + 1, intY - 1)-(Screen.ActiveForm.ScaleWidth, intY + intBANDHEIGHT), RGB(sngBlueCur, 0, 0), BF
    End If
   Case Else
    Screen.ActiveForm.Line (0, 0)-(sBarWidth, Screen.ActiveForm.Height), col1&, BF
  End Select
  sngBlueCur = sngBlueCur + sngBlueStep
 Next intY
End Sub
Public Sub AutoResize()
 Refresh
End Sub
Public Sub Refresh()
 Destroy
 sbActive = True
 oldMode% = Screen.ActiveForm.ScaleMode
 Screen.ActiveForm.ScaleMode = 3
 If sbFontName = "" Then sbFontName = "Arial"
 If sFillStyle = 0 Then sFillStyle = sbFilled
 If sAlign = 0 Then sAlign = sbAlignLeft
 If sbFontSize = 0 Then sbFontSize = 14
 Screen.ActiveForm.AutoRedraw = True
 If sFillStyle = sbFilled Then
  If sAlign = 1 Then
   Screen.ActiveForm.Line (0, 0)-(sBarWidth, Screen.ActiveForm.Height), sbBackColor, BF
  Else
   Screen.ActiveForm.Line (Screen.ActiveForm.ScaleWidth - sBarWidth, 0)-(Screen.ActiveForm.ScaleWidth, Screen.ActiveForm.Height), sbBackColor, BF
  End If
 ElseIf sFillStyle = sbSmooth Then
  Farbverlauf sbBackColor
 End If
 BckupCol& = Screen.ActiveForm.ForeColor
 Screen.ActiveForm.ForeColor = sbForeColor
 Dim hFont&, fontOld&
 hFont = CreateFont(-sbFontSize, 0, 90 * 10, 0, 700, False, False, 0, 1, 4, &H10, 2, 4, sbFontName)
 fontOld = SelectObject(Screen.ActiveForm.hdc, hFont)
 Screen.ActiveForm.ScaleMode = 3
 If sAlign = 1 Then
  TextOut Screen.ActiveForm.hdc, 2, Screen.ActiveForm.ScaleHeight - 8, sbText, Len(sbText)
 Else
  TextOut Screen.ActiveForm.hdc, (Screen.ActiveForm.ScaleWidth - sBarWidth) + 2, Screen.ActiveForm.ScaleHeight - 8, sbText, Len(sbText)
 End If
 SelectObject Screen.ActiveForm.hdc, fontOld
 DeleteObject hFont
 Screen.ActiveForm.ForeColor = BckupCol&
 Screen.ActiveForm.ScaleMode = oldMode%
End Sub
Public Sub RemoveTitleBar()
 SetWindowLong Screen.ActiveForm.hWnd, GWL_STYLE, GetWindowLong(Screen.ActiveForm.hWnd, GWL_STYLE) And Not WS_CAPTION
 oldMode% = Screen.ActiveForm.ScaleMode
 Screen.ActiveForm.ScaleMode = 1
 Screen.ActiveForm.Height = Screen.ActiveForm.Height - 300
 Screen.ActiveForm.ScaleMode = oldMode%
End Sub

