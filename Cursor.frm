VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   3240
   ClientLeft      =   1380
   ClientTop       =   1350
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   240
      Left            =   0
      Picture         =   "Cursor.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   1
      Top             =   3375
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.TextBox txtEDITOR 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3240
      Left            =   0
      MousePointer    =   2  'Cross
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5340
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' LICENSE AGREEMENT:
' Whatever you do, please notice that I will not HELD responsibility
' for any Illegal Act you might have upon my Creation.
' ==================================================================


Private Declare Function CreateCaret Lib _
"User32" (ByVal hwnd As Long, _
ByVal hBitmap As Long, ByVal nWidth _
As Long, ByVal nHeight As Long) As Long

Private Declare Function ShowCaret Lib _
"User32" (ByVal hwnd As Long) As Long

Private Declare Function GetFocus Lib _
"User32" () As Long

Private Declare Function SetCaretBlinkTime Lib "User32" _
(ByVal uMSeconds As Long) As Long

Private Sub Form_Load()

 Call SetCaretBlinkTime(900)
End Sub

Private Sub Form_Resize()
 txtEDITOR.Height = Me.Height
 txtEDITOR.Width = Me.Width
 h& = GetFocus&()
b& = Picture1.Picture
Call CreateCaret(h&, b&, 8, 10)
'handle, bitmap 0=none, width, height
X& = ShowCaret&(h&)
End Sub


Private Sub txteditor_Change()
h& = GetFocus&()
b& = Picture1.Picture
Call CreateCaret(h&, b&, 8, 10)
'handle, bitmap 0=none, width, height
X& = ShowCaret&(h&)
End Sub


