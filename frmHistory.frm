VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHistory 
   BorderStyle     =   0  'None
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHistory.frx":0000
   ScaleHeight     =   4635
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3120
      Left            =   315
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   660
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   5503
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmHistory.frx":28BB
   End
   Begin VB.Label lblClose 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   3900
      TabIndex        =   5
      Top             =   3915
      Width           =   1095
   End
   Begin VB.Label lblClear 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   2715
      TabIndex        =   4
      Top             =   3915
      Width           =   1095
   End
   Begin VB.Label lblOk 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   3915
      Width           =   1095
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   6180
      TabIndex        =   2
      Top             =   135
      Width           =   225
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "History"
      Height          =   195
      Left            =   195
      TabIndex        =   1
      Top             =   150
      Width           =   5910
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    On Error Resume Next
    RichTextBox1.LoadFile App.Path & "\History.dat"

End Sub

Private Sub lblClear_Click()

    On Error Resume Next
    Kill App.Path & "\History.dat"
    RichTextBox1.Text = ""

End Sub

Private Sub lblClose_Click()

    Unload Me

End Sub

Private Sub lblOk_Click()

    On Error Resume Next

    If RichTextBox1.SelText = "" Then

        Beep

    Else

        Dim i As Byte
        Dim strtemp As String, strtmp As String
        strtemp = RichTextBox1.SelText

        For i = 1 To 255

            strtmp = Mid(strtemp, i, 7)
        
            If strtmp = "_______" Then

                strtemp = Mid(strtemp, 2, i - 2)
                Exit For

            End If

        Next

        frmMain.Combo1.Text = strtemp
        frmMain.WebBrowser1.Navigate frmMain.Combo1.Text
        Call frmMain.History
        Unload Me

    End If

End Sub

Private Sub lblX_Click()

    Unload Me

End Sub

