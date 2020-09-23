Attribute VB_Name = "Module1"
Option Explicit

'For Progressbar
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Private Const WM_USER = &H400
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)

Public Sub ProgColor(PBR As ProgressBar, Backcolor As Long, Forecolor As Long)

    SendMessage PBR.hWnd, CCM_SETBKCOLOR, 0, ByVal Backcolor
    SendMessage PBR.hWnd, PBM_SETBARCOLOR, 0, ByVal Forecolor

End Sub

