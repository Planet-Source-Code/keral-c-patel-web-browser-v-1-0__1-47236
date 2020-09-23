VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnupopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuMplayer 
         Caption         =   "&MediaPlayer"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&ABOUT"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuAbout_Click()

    frmMain.WebBrowser1.Navigate App.Path & "\Welcome.htm"

End Sub

Private Sub mnuMplayer_Click()
If mnuMplayer.Checked = False Then
    mnuMplayer.Checked = True
    frmMplayer.Show
Else
    mnuMplayer.Checked = False
    Unload frmMplayer
End If
End Sub

