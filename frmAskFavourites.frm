VERSION 5.00
Begin VB.Form frmAskFavorites 
   BorderStyle     =   0  'None
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAskFavourites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAskFavourites.frx":000C
   ScaleHeight     =   1590
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   4875
      TabIndex        =   5
      Top             =   120
      Width           =   210
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "What You Want To Do in Favourites?"
      Height          =   195
      Left            =   1350
      TabIndex        =   4
      Top             =   705
      Width           =   2685
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Favourites:-"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   150
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   345
      Left            =   3990
      TabIndex        =   2
      Top             =   1095
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   345
      Left            =   2775
      TabIndex        =   1
      Top             =   1095
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Top             =   1095
      Width           =   1095
   End
End
Attribute VB_Name = "frmAskFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()

    frmFavorites.Show
    Unload Me

End Sub

Private Sub Label2_Click()

    frmViewFavorites.Show
    Unload Me

End Sub

Private Sub Label3_Click()

    Unload Me

End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&

End Sub

Private Sub Label6_Click()

    Unload Me

End Sub

