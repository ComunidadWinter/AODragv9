Attribute VB_Name = "modUrlDetection"
Option Explicit

Private Type NMHDR
hWndFrom As Long
idFrom As Long
code As Long
End Type

Private Type CHARRANGE
cpMin As Long
cpMax As Long
End Type

Private Type ENLINK
hdr As NMHDR
msg As Long
wParam As Long
lParam As Long
chrg As CHARRANGE
End Type

Private Type TEXTRANGE
chrg As CHARRANGE
lpstrText As String
End Type

'Used to change the window procedure which kick-starts the subclassing
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

'Used to call the default window procedure for the parent
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
ByVal lpPrevWndFunc As Long, _
ByVal hwnd As Long, _
ByVal msg As Long, _
ByVal wParam As Long, _
ByVal lParam As Long) As Long

'Used to set and retrieve various information
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long

'Used to copy... memory... from pointers
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
Destination As Any, _
Source As Any, _
ByVal Length As Long)

'Used to launch the URL in the user's default browser
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Const WM_NOTIFY = &H4E
Const EM_SETEVENTMASK = &H445
Const EM_GETEVENTMASK = &H43B
Const EM_GETTEXTRANGE = &H44B
Const EM_AUTOURLDETECT = &H45B
Const EN_LINK = &H70B

Const WM_LBUTTONDBLCLK = &H203
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_MOUSEMOVE = &H200
Const WM_RBUTTONDBLCLK = &H206
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_SETCURSOR = &H20

Const CFE_LINK = &H20
Const ENM_LINK = &H4000000
Const GWL_WNDPROC = (-4)
Const SW_SHOW = 5

Dim lOldProc As Long 'Old windowproc
Dim hWndRTB As Long 'hWnd of RTB
Dim hWndParent As Long 'hWnd of parent window

Public Sub DisableUrlDetect()

'Stop URL detection
SendMessage hWndRTB, EM_AUTOURLDETECT, 0, ByVal 0

End Sub

Public Sub EnableUrlDetect()

'Tell the RTB to start automatically detecting URLs
SendMessage hWndRTB, EM_AUTOURLDETECT, 1, ByVal 0

End Sub

Public Sub StartURLDetect(ByVal hWndTextbox As Long, ByVal hWndOwner As Long)
'Don't want to subclass twice!
If lOldProc = 0 Then
'Subclass!
lOldProc = SetWindowLong(hWndOwner, GWL_WNDPROC, AddressOf WndProc)

'Tell the RTB to inform us when stuff happens to URLs
SendMessage hWndTextbox, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(hWndTextbox, EM_GETEVENTMASK, 0, 0)

hWndParent = hWndOwner
hWndRTB = hWndTextbox

End If

End Sub

Public Sub StopURLDetect()
'Don't want to unsubclass a non-subclassed window
If lOldProc Then
'Stop URL detection
SendMessage hWndRTB, EM_AUTOURLDETECT, 0, ByVal 0
'Reset the window procedure (stop the subclassing)
SetWindowLong hWndParent, GWL_WNDPROC, lOldProc
'Set this to 0 so we can subclass again in future
lOldProc = 0
End If

End Sub

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim uHead As NMHDR
Dim eLink As ENLINK
Dim eText As TEXTRANGE
Dim sText As String
Dim lLen As Long

'Which message?
Select Case uMsg
Case WM_NOTIFY

'Ooo! A notify message! Something exciting must be happening...
'Copy the notification header into our structure from the pointer
CopyMemory uHead, ByVal lParam, Len(uHead)

'Peek inside the structure
If (uHead.hWndFrom = hWndRTB) And (uHead.code = EN_LINK) Then

'Yay! Some kind of kinky linky message.
'Now that we know its a link message, we can copy the whole ENLINK structure
'into our structure
CopyMemory eLink, ByVal lParam, Len(eLink)

'What kind of message?
Select Case eLink.msg
Case WM_LBUTTONDOWN 'WM_LBUTTONDBLCLK
'Double clicked the link!

'Set up out TEXTRANGE struct
eText.chrg.cpMin = eLink.chrg.cpMin
eText.chrg.cpMax = eLink.chrg.cpMax
eText.lpstrText = Space$(1024)

'Tell the RTB to fill out our TEXTRANGE with the text
lLen = SendMessage(hWndRTB, EM_GETTEXTRANGE, 0, eText)

'Trim the text
sText = Left$(eText.lpstrText, lLen)

'Launch the browser
ShellExecute hWndParent, vbNullString, sText, vbNullString, vbNullString, SW_SHOW

'Other miscellaneous messages
Case WM_LBUTTONDOWN

Case WM_LBUTTONUP

Case WM_RBUTTONDBLCLK

Case WM_RBUTTONDOWN

Case WM_RBUTTONUP

Case WM_SETCURSOR

End Select

End If

End Select

'Call the stored window procedure to let it handle all the messages
WndProc = CallWindowProc(lOldProc, hwnd, uMsg, wParam, lParam)
End Function



