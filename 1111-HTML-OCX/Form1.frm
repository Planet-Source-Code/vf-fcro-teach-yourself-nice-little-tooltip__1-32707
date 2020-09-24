VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Author:VANJA FUCKAR,EMAIL:INGA@VIP.HR"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Just CLICK!"
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "TOOLTIP FROM HTML HELP CONTROL!!!"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const HH_DISPLAY_TEXT_POPUP = &HE
Private Declare Function HtmlHelp Lib "HHCtrl.ocx" _
        Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, _
        ByVal uCommand As Long, dwData As Any) As Long


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
     x As Long
     y As Long
End Type

Private Type HH_POPUP
        cbStruct As Long
        hinst As Long
        idString As Long
        pszText As Long
        pt As POINTAPI
        clrForeground As Long
        clrBackground As Long
        rcMargins As RECT
        pzsFont As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Private Sub Command1_Click()
Dim text1 As String
text1 = "What You Think About That Nice Little ToolTip??? HA?" & Chr(CByte(0))
Dim AnsiText() As Byte
AnsiText = StrConv(text1, vbFromUnicode)

Dim POSITION As POINTAPI
Dim PTHELP As POINTAPI
Dim RECTHELP As RECT

Dim FNT As String
FNT = "Arial Black" & Chr(CByte(0))
Dim FNT1() As Byte
FNT1 = StrConv(FNT, vbFromUnicode)

GetCursorPos POSITION

With PTHELP
.x = POSITION.x
.y = POSITION.y
End With

With RECTHELP
.Left = 10
.Top = 10
.Bottom = 10
.Right = 10
End With

Dim HHTEXT As HH_POPUP

With HHTEXT
.cbStruct = Len(HHTEXT)
.hinst = 0
.idString = 0
.pszText = VarPtr(AnsiText(0))
.clrBackground = &HC0FFFF
.clrForeground = &HFF3333
.rcMargins = RECTHELP
.pt = PTHELP
.pzsFont = VarPtr(FNT1(0))
End With

Call HtmlHelp(0&, 0&, HH_DISPLAY_TEXT_POPUP, HHTEXT)
End Sub

