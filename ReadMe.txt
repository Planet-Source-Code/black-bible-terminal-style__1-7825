If you can't access this project, then use this code:

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

PlACE this code on FORM_LOAD:  Call SetCaretBlinkTime(900)

create a textbox and name it to txtEDITOR and place this code into
txtEDITOR_CHANGE: 
h& = GetFocus&()
b& = Picture1.Picture
Call CreateCaret(h&, b&, 8, 10)
'handle, bitmap 0=none, width, height
X& = ShowCaret&(h&)


now, create a picturebox and place a little green cursor. ;-)
you can do this in paint or any paint program.



