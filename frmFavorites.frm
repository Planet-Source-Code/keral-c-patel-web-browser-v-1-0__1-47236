VERSION 5.00
Begin VB.Form frmFavorites 
   BorderStyle     =   0  'None
   Caption         =   "Favorites..."
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFavorites.frx":0000
   ScaleHeight     =   1590
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4845
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add To Favorites....."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4560
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim num As Integer

Private Sub Form_Load()

    On Error Resume Next
    num = GetSetting(App.EXEName, "Fav", "Number")
    Text1 = frmMain.Combo1.Text

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&

End Sub

Private Sub Label2_Click()

    On Error Resume Next
    'to increase the number
    SaveSetting App.EXEName, "Fav", "Number", CStr(num + 1)
    'to add address
    SaveSetting App.EXEName, "Fav", "add" & (num + 1), Text1.Text
    Unload Me

End Sub

Private Sub Label3_Click()

    Unload Me

End Sub

