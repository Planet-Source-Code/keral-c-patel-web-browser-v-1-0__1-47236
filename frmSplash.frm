VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5025
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3390
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   315
      Top             =   3315
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   135
      TabIndex        =   0
      Top             =   3030
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   -315
      Top             =   3330
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By Keral.C.Patel."
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   2760
      Width           =   2010
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    ProgColor ProgressBar1, vbWhite, vbRed
    Load frmMain

End Sub

Private Sub Timer1_Timer()

    If ProgressBar1.Value = 100 Then

        frmMain.Visible = True
        frmMain.Timer2.Enabled = True
        Timer1.Enabled = False
        Unload Me

    Else

        ProgressBar1.Value = ProgressBar1.Value + 2

    End If

End Sub

