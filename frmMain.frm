VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9390
   ClientLeft      =   3195
   ClientTop       =   1965
   ClientWidth     =   12450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   830
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1710
      Top             =   7725
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   975
      Top             =   7695
   End
   Begin VB.CheckBox chkFavorites 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Favorites"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   6090
      Picture         =   "frmMain.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   510
      Width           =   855
   End
   Begin VB.CheckBox chkHistory 
      BackColor       =   &H00FFFFFF&
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   5250
      Picture         =   "frmMain.frx":1A88
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   510
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   210
      TabIndex        =   10
      Text            =   "about:blank"
      Top             =   1395
      Width           =   11055
   End
   Begin VB.CheckBox chkGo 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   11370
      Picture         =   "frmMain.frx":1F3E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1395
      Width           =   495
   End
   Begin VB.CheckBox chkHome 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   4410
      Picture         =   "frmMain.frx":2250
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   510
      Width           =   855
   End
   Begin VB.CheckBox chkSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   3570
      Picture         =   "frmMain.frx":2B1A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   510
      Width           =   855
   End
   Begin VB.CheckBox chkRefresh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   2730
      Picture         =   "frmMain.frx":300C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   510
      Width           =   855
   End
   Begin VB.CheckBox chkStop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1890
      Picture         =   "frmMain.frx":342A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   510
      Width           =   855
   End
   Begin VB.CheckBox chkForward 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Forward"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1050
      Picture         =   "frmMain.frx":38E0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   510
      Width           =   855
   End
   Begin VB.CheckBox chkBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   210
      Picture         =   "frmMain.frx":4142
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   510
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   7290
      TabIndex        =   1
      Top             =   1020
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5355
      Left            =   210
      TabIndex        =   0
      Top             =   1755
      Width           =   9285
      ExtentX         =   16378
      ExtentY         =   9446
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Image imgskin 
      Height          =   375
      Left            =   4065
      Picture         =   "frmMain.frx":49A4
      Stretch         =   -1  'True
      Top             =   7410
      Width           =   420
   End
   Begin VB.Image imgMin 
      Height          =   255
      Left            =   9870
      Picture         =   "frmMain.frx":521A
      Top             =   7800
      Width           =   255
   End
   Begin VB.Image imgMax 
      Height          =   255
      Left            =   10125
      Picture         =   "frmMain.frx":55D2
      Top             =   7800
      Width           =   255
   End
   Begin VB.Image imgClose 
      Height          =   255
      Left            =   10380
      Picture         =   "frmMain.frx":599F
      Top             =   7800
      Width           =   255
   End
   Begin VB.Label lblBrowser 
      BackStyle       =   0  'Transparent
      Caption         =   "WebBrowser"
      Height          =   195
      Left            =   330
      TabIndex        =   13
      Top             =   225
      Width           =   915
   End
   Begin VB.Image imgHeader 
      Height          =   255
      Left            =   225
      Picture         =   "frmMain.frx":5D6D
      Top             =   180
      Width           =   14010
   End
   Begin VB.Image Image3to4 
      Height          =   5415
      Left            =   2790
      Picture         =   "frmMain.frx":703C
      Stretch         =   -1  'True
      Top             =   7665
      Width           =   180
   End
   Begin VB.Image Image2to3 
      Height          =   180
      Left            =   1095
      Picture         =   "frmMain.frx":7AB3
      Stretch         =   -1  'True
      Top             =   8310
      Width           =   6735
   End
   Begin VB.Image Image1to2 
      Height          =   5925
      Left            =   3105
      Picture         =   "frmMain.frx":7EE1
      Stretch         =   -1  'True
      Top             =   7380
      Width           =   150
   End
   Begin VB.Image Image1to4 
      Height          =   165
      Left            =   1230
      Picture         =   "frmMain.frx":87BF
      Stretch         =   -1  'True
      Top             =   8100
      Width           =   6840
   End
   Begin VB.Image Image4 
      Height          =   120
      Left            =   5520
      Picture         =   "frmMain.frx":8BD4
      Top             =   7365
      Width           =   165
   End
   Begin VB.Image Image3 
      Height          =   150
      Left            =   600
      Picture         =   "frmMain.frx":8F2A
      Top             =   8250
      Width           =   165
   End
   Begin VB.Image Image2 
      Height          =   150
      Left            =   45
      Picture         =   "frmMain.frx":929B
      Top             =   8295
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   120
      Left            =   7320
      Picture         =   "frmMain.frx":95F3
      Top             =   7335
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BlueSoft®"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   465
      Left            =   7590
      TabIndex        =   9
      Top             =   540
      Width           =   1770
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkHistory_Click()

    On Error Resume Next
    Dim xx As Byte
    xx = xx + chkHistory.Value
    chkHistory.Value = 0

    If xx = 1 Then

        frmHistory.Show
        
    End If

End Sub

Private Sub chkFavorites_Click()

    On Error Resume Next
    Dim xx As Byte
    xx = xx + chkFavorites.Value
    chkFavorites.Value = 0

    If xx = 1 Then

        frmAskFavorites.Show

    End If

End Sub

Private Sub chkForward_Click()

    On Error Resume Next
    Dim xx As Byte
    xx = xx + chkForward.Value
    chkForward.Value = 0

    If xx = 1 Then

        On Error Resume Next
        WebBrowser1.GoForward

    End If

End Sub

Private Sub chkBack_Click()

    On Error Resume Next
    Dim xx As Byte
    xx = xx + chkBack.Value
    chkBack.Value = 0

    If xx = 1 Then

        On Error Resume Next
        WebBrowser1.GoBack

    End If

End Sub

Private Sub chkStop_Click()

    On Error Resume Next
    Dim xx As Byte
    xx = xx + chkStop.Value
    chkStop.Value = 0

    If xx = 1 Then

        On Error Resume Next
        WebBrowser1.Stop

    End If

End Sub

Private Sub chkRefresh_Click()

    On Error Resume Next
    Dim xx As Byte
    xx = xx + chkRefresh.Value
    chkRefresh.Value = 0

    If xx = 1 Then

        On Error Resume Next
        WebBrowser1.Refresh

    End If

End Sub

Private Sub chkHome_Click()

    On Error Resume Next
    Dim xx As Byte
    xx = xx + chkHome.Value
    chkHome.Value = 0

    If xx = 1 Then

        On Error Resume Next
        WebBrowser1.GoHome

    End If

End Sub

Private Sub chkSearch_Click()

    On Error Resume Next
    Dim xx As Byte
    xx = xx + chkSearch.Value
    chkSearch.Value = 0

    If xx = 1 Then

        WebBrowser1.GoSearch

    End If

End Sub

Private Sub chkGo_Click()

    On Error Resume Next
    Dim xx As Byte
    xx = xx + chkGo.Value
    chkGo.Value = 0

    If xx = 1 Then

        On Error Resume Next
        WebBrowser1.Navigate (Combo1.Text)
        Call History

    End If

End Sub

Private Sub Form_Load()

    On Error Resume Next
    Call setskin
    WebBrowser1.Navigate App.Path & "\Welcome.htm"
    ProgColor ProgressBar1, vbWhite, vbRed

End Sub

Private Sub lblResize_Click()

End Sub

Private Sub imgClose_Click()

    Unload Me

End Sub

Private Sub imgHeader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then

        ReleaseCapture
        SendMessage Me.hWnd, &HA1, 2, 0&

    Else

        PopupMenu frmMenu.mnupopup

    End If

End Sub

Private Sub imgMax_Click()

    If Me.WindowState = 0 Then

        Me.WindowState = 2

    Else

        Me.WindowState = 0

    End If

End Sub

Private Sub imgMin_Click()

    Me.WindowState = 1

End Sub

Private Sub Timer1_Timer()

    If lblBrowser.Left >= 50 Then

        lblBrowser.Left = lblBrowser.Left - 5

    Else

        Timer1.Enabled = False
        Timer2.Enabled = True

    End If

End Sub

Private Sub Timer2_Timer()

    If lblBrowser.Left <= Me.ScaleWidth - 165 Then

        lblBrowser.Left = lblBrowser.Left + 5

    Else

        Timer2.Enabled = False
        Timer1.Enabled = True

    End If

End Sub

Private Sub WebBrowser1_ProgressChange(ByVal _
   Progress As Long, ByVal ProgressMax As Long)

    On Error Resume Next
    ProgressBar1.Max = ProgressMax
    ProgressBar1.Value = Progress
    ProgressBar1.Refresh

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then

        WebBrowser1.Navigate (Combo1)

    End If

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Combo1.Width = Me.ScaleWidth - 70
    chkGo.Left = Combo1.Left + Combo1.Width + 10
    WebBrowser1.Width = Me.ScaleWidth - 30
    WebBrowser1.Height = Me.ScaleHeight - (Combo1.Top + Combo1.Height) - 20
    Call setskin

End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As _
   Object, URL As Variant)

    Combo1.Text = WebBrowser1.LocationURL

End Sub

Private Sub setskin()

    On Error Resume Next
    '*****************************************************************************
    'For setting the skin
    'Header
    imgHeader.Top = 12
    imgHeader.Width = Me.ScaleWidth - 100
    'Other Borders
    Image1.Left = 0
    Image1.Top = 0
    Image2.Left = 0
    Image2.Top = Me.ScaleHeight - Image2.Height
    Image3.Left = Me.ScaleWidth - Image3.Width
    Image3.Top = Me.ScaleHeight - Image3.Height
    Image4.Left = Me.ScaleWidth - Image4.Width
    Image4.Top = 0
    'for image1to2
    Image1to2.Left = 0
    Image1to2.Top = Image1.Height
    Image1to2.Height = Me.ScaleHeight - (Image1.Height + Image2.Height)
    'for image2to3
    Image2to3.Left = Image2.Width
    Image2to3.Top = Me.ScaleHeight - Image2to3.Height
    Image2to3.Width = Me.ScaleWidth - (Image2.Width + Image3.Width)
    'for image3to4
    Image3to4.Top = Image4.Height
    Image3to4.Left = Me.ScaleWidth - Image3to4.Width
    Image3to4.Height = Me.ScaleHeight - (Image3.Height + Image4.Height)
    'for image1to4
    Image1to4.Top = 0
    Image1to4.Left = Image1.Width
    Image1to4.Width = Me.ScaleWidth - (Image1.Width + Image4.Width)
    'For Controlbox buttons
    imgClose.Left = Me.ScaleWidth - 35
    imgClose.Top = 12
    imgMax.Left = Me.ScaleWidth - 55
    imgMax.Top = 12
    imgMin.Left = Me.ScaleWidth - 75
    imgMin.Top = 12
    'For placing webbrowser in front of all controls
    WebBrowser1.ZOrder 0
    '*****************************************************************************

End Sub

Public Sub History()

    On Error Resume Next
    Dim dteNow As Date, dteToday As Date
    dteNow = Time
    dteToday = Date
    Open App.Path & "\History.dat" For Append As #1
    Write #1, Combo1.Text & "_______" & dteNow & "--On--" & dteToday
    Close #1

End Sub

