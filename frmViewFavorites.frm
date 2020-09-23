VERSION 5.00
Begin VB.Form frmViewFavorites 
   BorderStyle     =   0  'None
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmViewFavorites.frx":0000
   ScaleHeight     =   4800
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3420
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2685
   End
   Begin VB.Label lblGo 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   960
      TabIndex        =   3
      Top             =   4125
      Width           =   1125
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "View Favourites"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   210
      TabIndex        =   2
      Top             =   150
      Width           =   2400
   End
   Begin VB.Label lblClose 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   2730
      TabIndex        =   1
      Top             =   135
      Width           =   240
   End
End
Attribute VB_Name = "frmViewFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    On Error Resume Next
    Dim num As String
    Dim i As Integer
    Dim strtemp As String
    num = GetSetting(App.EXEName, "Fav", "Number")

    If num = "0" Then Exit Sub

    DoEvents

    If CInt(num) > 0 Then

        For i = 1 To num

            strtemp = GetSetting(App.EXEName, "Fav", "add" & i)
            List1.AddItem strtemp

        Next

    End If

End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&

End Sub

Private Sub lblClose_Click()

    Unload Me

End Sub

Private Sub lblGo_Click()

    frmMain.Combo1.Text = List1.Text
    DoEvents
    frmMain.WebBrowser1.Navigate (frmMain.Combo1.Text)
    Unload Me

End Sub

