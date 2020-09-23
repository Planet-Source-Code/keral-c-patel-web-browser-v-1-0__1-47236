VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMplayer 
   BorderStyle     =   0  'None
   Caption         =   "Media Player"
   ClientHeight    =   2460
   ClientLeft      =   4545
   ClientTop       =   4530
   ClientWidth     =   6270
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   Picture         =   "frmMPlayer.frx":0000
   ScaleHeight     =   2460
   ScaleWidth      =   6270
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   585
      Top             =   1080
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   135
      Top             =   1890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   915
      TabIndex        =   0
      Top             =   2415
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By:- Keral.C.Patel."
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   2175
      Width           =   2100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6250
      TabIndex        =   5
      Top             =   750
      Width           =   60
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   345
      Index           =   3
      Left            =   4080
      TabIndex        =   4
      Top             =   1530
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   345
      Index           =   2
      Left            =   3165
      TabIndex        =   3
      Top             =   1530
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   345
      Index           =   1
      Left            =   2235
      TabIndex        =   2
      Top             =   1530
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   345
      Index           =   0
      Left            =   1305
      TabIndex        =   1
      Top             =   1530
      Width           =   930
   End
End
Attribute VB_Name = "frmMplayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Const RGN_COPY = 5
Private ResultRegion As Long
Dim strfilename As String 'for storing the filename of the song

Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long

    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)
    ReDim PolyPoints(0 To 187)

    For Counter = 0 To 187

        PolyPoints(Counter).X = GP0X(Counter) * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX
        PolyPoints(Counter).Y = GP0Y(Counter) * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY

    Next Counter

    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 188, 1)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 26)

    For Counter = 0 To 26

        PolyPoints(Counter).X = GP1X(Counter) * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX
        PolyPoints(Counter).Y = GP1Y(Counter) * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY

    Next Counter

    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 27, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion

End Function

Private Function GP0X(Number As Integer) As Integer

    Select Case Number

        Case 0
            GP0X = 104

        Case 1
            GP0X = 314

        Case 2
            GP0X = 315

        Case 3
            GP0X = 316

        Case 4
            GP0X = 318

        Case 5
            GP0X = 319

        Case 6
            GP0X = 323

        Case 7
            GP0X = 324

        Case 8
            GP0X = 327

        Case 9
            GP0X = 328

        Case 10
            GP0X = 332

        Case 11
            GP0X = 333

        Case 12
            GP0X = 337

        Case 13
            GP0X = 338

        Case 14
            GP0X = 341

        Case 15
            GP0X = 342

        Case 16
            GP0X = 346

        Case 17
            GP0X = 347

        Case 18
            GP0X = 351

        Case 19
            GP0X = 352

        Case 20
            GP0X = 355

        Case 21
            GP0X = 356

        Case 22
            GP0X = 360

        Case 23
            GP0X = 361

        Case 24
            GP0X = 365

        Case 25
            GP0X = 366

        Case 26
            GP0X = 369

        Case 27
            GP0X = 370

        Case 28
            GP0X = 374

        Case 29
            GP0X = 375

        Case 30
            GP0X = 379

        Case 31
            GP0X = 380

        Case 32
            GP0X = 383

        Case 33
            GP0X = 384

        Case 34
            GP0X = 388

        Case 35
            GP0X = 389

        Case 36
            GP0X = 393

        Case 37
            GP0X = 394

        Case 38
            GP0X = 398

        Case 39
            GP0X = 399

        Case 40
            GP0X = 402

        Case 41
            GP0X = 403

        Case 42
            GP0X = 407

        Case 43
            GP0X = 408

        Case 44
            GP0X = 412

        Case 45
            GP0X = 413

        Case 46
            GP0X = 416

        Case 47
            GP0X = 417

        Case 48
            GP0X = 418

        Case 49
            GP0X = 418

        Case 50
            GP0X = 414

        Case 51
            GP0X = 413

        Case 52
            GP0X = 410

        Case 53
            GP0X = 409

        Case 54
            GP0X = 405

        Case 55
            GP0X = 404

        Case 56
            GP0X = 401

        Case 57
            GP0X = 400

        Case 58
            GP0X = 396

        Case 59
            GP0X = 395

        Case 60
            GP0X = 392

        Case 61
            GP0X = 391

        Case 62
            GP0X = 387

        Case 63
            GP0X = 386

        Case 64
            GP0X = 383

        Case 65
            GP0X = 382

        Case 66
            GP0X = 378

        Case 67
            GP0X = 377

        Case 68
            GP0X = 374

        Case 69
            GP0X = 373

        Case 70
            GP0X = 369

        Case 71
            GP0X = 368

        Case 72
            GP0X = 365

        Case 73
            GP0X = 364

        Case 74
            GP0X = 360

        Case 75
            GP0X = 359

        Case 76
            GP0X = 356

        Case 77
            GP0X = 355

        Case 78
            GP0X = 351

        Case 79
            GP0X = 350

        Case 80
            GP0X = 347

        Case 81
            GP0X = 346

        Case 82
            GP0X = 342

        Case 83
            GP0X = 341

        Case 84
            GP0X = 338

        Case 85
            GP0X = 337

        Case 86
            GP0X = 333

        Case 87
            GP0X = 332

        Case 88
            GP0X = 329

        Case 89
            GP0X = 328

        Case 90
            GP0X = 324

        Case 91
            GP0X = 323

        Case 92
            GP0X = 320

        Case 93
            GP0X = 319

        Case 94
            GP0X = 315

        Case 95
            GP0X = 312

        Case 96
            GP0X = 104

        Case 97
            GP0X = 101

        Case 98
            GP0X = 100

        Case 99
            GP0X = 96

        Case 100
            GP0X = 95

        Case 101
            GP0X = 92

        Case 102
            GP0X = 91

        Case 103
            GP0X = 87

        Case 104
            GP0X = 86

        Case 105
            GP0X = 82

        Case 106
            GP0X = 81

        Case 107
            GP0X = 78

        Case 108
            GP0X = 77

        Case 109
            GP0X = 73

        Case 110
            GP0X = 72

        Case 111
            GP0X = 68

        Case 112
            GP0X = 67

        Case 113
            GP0X = 63

        Case 114
            GP0X = 62

        Case 115
            GP0X = 59

        Case 116
            GP0X = 58

        Case 117
            GP0X = 54

        Case 118
            GP0X = 53

        Case 119
            GP0X = 49

        Case 120
            GP0X = 48

        Case 121
            GP0X = 45

        Case 122
            GP0X = 44

        Case 123
            GP0X = 40

        Case 124
            GP0X = 39

        Case 125
            GP0X = 35

        Case 126
            GP0X = 34

        Case 127
            GP0X = 31

        Case 128
            GP0X = 30

        Case 129
            GP0X = 26

        Case 130
            GP0X = 25

        Case 131
            GP0X = 21

        Case 132
            GP0X = 20

        Case 133
            GP0X = 16

        Case 134
            GP0X = 15

        Case 135
            GP0X = 12

        Case 136
            GP0X = 11

        Case 137
            GP0X = 7

        Case 138
            GP0X = 6

        Case 139
            GP0X = 2

        Case 140
            GP0X = 1

        Case 141
            GP0X = 0

        Case 142
            GP0X = 1

        Case 143
            GP0X = 5

        Case 144
            GP0X = 6

        Case 145
            GP0X = 9

        Case 146
            GP0X = 10

        Case 147
            GP0X = 14

        Case 148
            GP0X = 15

        Case 149
            GP0X = 18

        Case 150
            GP0X = 19

        Case 151
            GP0X = 22

        Case 152
            GP0X = 23

        Case 153
            GP0X = 27

        Case 154
            GP0X = 28

        Case 155
            GP0X = 31

        Case 156
            GP0X = 32

        Case 157
            GP0X = 36

        Case 158
            GP0X = 37

        Case 159
            GP0X = 40

        Case 160
            GP0X = 41

        Case 161
            GP0X = 44

        Case 162
            GP0X = 45

        Case 163
            GP0X = 49

        Case 164
            GP0X = 50

        Case 165
            GP0X = 53

        Case 166
            GP0X = 54

        Case 167
            GP0X = 58

        Case 168
            GP0X = 59

        Case 169
            GP0X = 62

        Case 170
            GP0X = 63

        Case 171
            GP0X = 66

        Case 172
            GP0X = 67

        Case 173
            GP0X = 71

        Case 174
            GP0X = 72

        Case 175
            GP0X = 75

        Case 176
            GP0X = 76

        Case 177
            GP0X = 80

        Case 178
            GP0X = 81

        Case 179
            GP0X = 84

        Case 180
            GP0X = 85

        Case 181
            GP0X = 88

        Case 182
            GP0X = 89

        Case 183
            GP0X = 93

        Case 184
            GP0X = 94

        Case 185
            GP0X = 97

        Case 186
            GP0X = 98

        Case 187
            GP0X = 102

    End Select

End Function

Private Function GP0Y(Number As Integer) As Integer

    Select Case Number

        Case 0
            GP0Y = 0

        Case 1
            GP0Y = 0

        Case 2
            GP0Y = 1

        Case 3
            GP0Y = 1

        Case 4
            GP0Y = 3

        Case 5
            GP0Y = 3

        Case 6
            GP0Y = 7

        Case 7
            GP0Y = 7

        Case 8
            GP0Y = 10

        Case 9
            GP0Y = 10

        Case 10
            GP0Y = 14

        Case 11
            GP0Y = 14

        Case 12
            GP0Y = 18

        Case 13
            GP0Y = 18

        Case 14
            GP0Y = 21

        Case 15
            GP0Y = 21

        Case 16
            GP0Y = 25

        Case 17
            GP0Y = 25

        Case 18
            GP0Y = 29

        Case 19
            GP0Y = 29

        Case 20
            GP0Y = 32

        Case 21
            GP0Y = 32

        Case 22
            GP0Y = 36

        Case 23
            GP0Y = 36

        Case 24
            GP0Y = 40

        Case 25
            GP0Y = 40

        Case 26
            GP0Y = 43

        Case 27
            GP0Y = 43

        Case 28
            GP0Y = 47

        Case 29
            GP0Y = 47

        Case 30
            GP0Y = 51

        Case 31
            GP0Y = 51

        Case 32
            GP0Y = 54

        Case 33
            GP0Y = 54

        Case 34
            GP0Y = 58

        Case 35
            GP0Y = 58

        Case 36
            GP0Y = 62

        Case 37
            GP0Y = 62

        Case 38
            GP0Y = 66

        Case 39
            GP0Y = 66

        Case 40
            GP0Y = 69

        Case 41
            GP0Y = 69

        Case 42
            GP0Y = 73

        Case 43
            GP0Y = 73

        Case 44
            GP0Y = 77

        Case 45
            GP0Y = 77

        Case 46
            GP0Y = 80

        Case 47
            GP0Y = 80

        Case 48
            GP0Y = 81

        Case 49
            GP0Y = 82

        Case 50
            GP0Y = 86

        Case 51
            GP0Y = 86

        Case 52
            GP0Y = 89

        Case 53
            GP0Y = 89

        Case 54
            GP0Y = 93

        Case 55
            GP0Y = 93

        Case 56
            GP0Y = 96

        Case 57
            GP0Y = 96

        Case 58
            GP0Y = 100

        Case 59
            GP0Y = 100

        Case 60
            GP0Y = 103

        Case 61
            GP0Y = 103

        Case 62
            GP0Y = 107

        Case 63
            GP0Y = 107

        Case 64
            GP0Y = 110

        Case 65
            GP0Y = 110

        Case 66
            GP0Y = 114

        Case 67
            GP0Y = 114

        Case 68
            GP0Y = 117

        Case 69
            GP0Y = 117

        Case 70
            GP0Y = 121

        Case 71
            GP0Y = 121

        Case 72
            GP0Y = 124

        Case 73
            GP0Y = 124

        Case 74
            GP0Y = 128

        Case 75
            GP0Y = 128

        Case 76
            GP0Y = 131

        Case 77
            GP0Y = 131

        Case 78
            GP0Y = 135

        Case 79
            GP0Y = 135

        Case 80
            GP0Y = 138

        Case 81
            GP0Y = 138

        Case 82
            GP0Y = 142

        Case 83
            GP0Y = 142

        Case 84
            GP0Y = 145

        Case 85
            GP0Y = 145

        Case 86
            GP0Y = 149

        Case 87
            GP0Y = 149

        Case 88
            GP0Y = 152

        Case 89
            GP0Y = 152

        Case 90
            GP0Y = 156

        Case 91
            GP0Y = 156

        Case 92
            GP0Y = 159

        Case 93
            GP0Y = 159

        Case 94
            GP0Y = 163

        Case 95
            GP0Y = 164

        Case 96
            GP0Y = 164

        Case 97
            GP0Y = 161

        Case 98
            GP0Y = 161

        Case 99
            GP0Y = 157

        Case 100
            GP0Y = 157

        Case 101
            GP0Y = 154

        Case 102
            GP0Y = 154

        Case 103
            GP0Y = 150

        Case 104
            GP0Y = 150

        Case 105
            GP0Y = 146

        Case 106
            GP0Y = 146

        Case 107
            GP0Y = 143

        Case 108
            GP0Y = 143

        Case 109
            GP0Y = 139

        Case 110
            GP0Y = 139

        Case 111
            GP0Y = 135

        Case 112
            GP0Y = 135

        Case 113
            GP0Y = 131

        Case 114
            GP0Y = 131

        Case 115
            GP0Y = 128

        Case 116
            GP0Y = 128

        Case 117
            GP0Y = 124

        Case 118
            GP0Y = 124

        Case 119
            GP0Y = 120

        Case 120
            GP0Y = 120

        Case 121
            GP0Y = 117

        Case 122
            GP0Y = 117

        Case 123
            GP0Y = 113

        Case 124
            GP0Y = 113

        Case 125
            GP0Y = 109

        Case 126
            GP0Y = 109

        Case 127
            GP0Y = 106

        Case 128
            GP0Y = 106

        Case 129
            GP0Y = 102

        Case 130
            GP0Y = 102

        Case 131
            GP0Y = 98

        Case 132
            GP0Y = 98

        Case 133
            GP0Y = 94

        Case 134
            GP0Y = 94

        Case 135
            GP0Y = 91

        Case 136
            GP0Y = 91

        Case 137
            GP0Y = 87

        Case 138
            GP0Y = 87

        Case 139
            GP0Y = 83

        Case 140
            GP0Y = 83

        Case 141
            GP0Y = 80

        Case 142
            GP0Y = 80

        Case 143
            GP0Y = 76

        Case 144
            GP0Y = 76

        Case 145
            GP0Y = 73

        Case 146
            GP0Y = 73

        Case 147
            GP0Y = 69

        Case 148
            GP0Y = 69

        Case 149
            GP0Y = 66

        Case 150
            GP0Y = 66

        Case 151
            GP0Y = 63

        Case 152
            GP0Y = 63

        Case 153
            GP0Y = 59

        Case 154
            GP0Y = 59

        Case 155
            GP0Y = 56

        Case 156
            GP0Y = 56

        Case 157
            GP0Y = 52

        Case 158
            GP0Y = 52

        Case 159
            GP0Y = 49

        Case 160
            GP0Y = 49

        Case 161
            GP0Y = 46

        Case 162
            GP0Y = 46

        Case 163
            GP0Y = 42

        Case 164
            GP0Y = 42

        Case 165
            GP0Y = 39

        Case 166
            GP0Y = 39

        Case 167
            GP0Y = 35

        Case 168
            GP0Y = 35

        Case 169
            GP0Y = 32

        Case 170
            GP0Y = 32

        Case 171
            GP0Y = 29

        Case 172
            GP0Y = 29

        Case 173
            GP0Y = 25

        Case 174
            GP0Y = 25

        Case 175
            GP0Y = 22

        Case 176
            GP0Y = 22

        Case 177
            GP0Y = 18

        Case 178
            GP0Y = 18

        Case 179
            GP0Y = 15

        Case 180
            GP0Y = 15

        Case 181
            GP0Y = 12

        Case 182
            GP0Y = 12

        Case 183
            GP0Y = 8

        Case 184
            GP0Y = 8

        Case 185
            GP0Y = 5

        Case 186
            GP0Y = 5

        Case 187
            GP0Y = 1

    End Select

End Function

Private Function GP1X(Number As Integer) As Integer

    Select Case Number

        Case 0
            GP1X = 389

        Case 1
            GP1X = 390

        Case 2
            GP1X = 391

        Case 3
            GP1X = 395

        Case 4
            GP1X = 396

        Case 5
            GP1X = 399

        Case 6
            GP1X = 400

        Case 7
            GP1X = 404

        Case 8
            GP1X = 405

        Case 9
            GP1X = 409

        Case 10
            GP1X = 410

        Case 11
            GP1X = 413

        Case 12
            GP1X = 414

        Case 13
            GP1X = 415

        Case 14
            GP1X = 415

        Case 15
            GP1X = 411

        Case 16
            GP1X = 410

        Case 17
            GP1X = 407

        Case 18
            GP1X = 406

        Case 19
            GP1X = 402

        Case 20
            GP1X = 401

        Case 21
            GP1X = 398

        Case 22
            GP1X = 397

        Case 23
            GP1X = 393

        Case 24
            GP1X = 392

        Case 25
            GP1X = 389

        Case 26
            GP1X = 388

    End Select

End Function

Private Function GP1Y(Number As Integer) As Integer

    Select Case Number

        Case 0
            GP1Y = 61

        Case 1
            GP1Y = 62

        Case 2
            GP1Y = 62

        Case 3
            GP1Y = 66

        Case 4
            GP1Y = 66

        Case 5
            GP1Y = 69

        Case 6
            GP1Y = 69

        Case 7
            GP1Y = 73

        Case 8
            GP1Y = 73

        Case 9
            GP1Y = 77

        Case 10
            GP1Y = 77

        Case 11
            GP1Y = 80

        Case 12
            GP1Y = 80

        Case 13
            GP1Y = 81

        Case 14
            GP1Y = 82

        Case 15
            GP1Y = 86

        Case 16
            GP1Y = 86

        Case 17
            GP1Y = 89

        Case 18
            GP1Y = 89

        Case 19
            GP1Y = 93

        Case 20
            GP1Y = 93

        Case 21
            GP1Y = 96

        Case 22
            GP1Y = 96

        Case 23
            GP1Y = 100

        Case 24
            GP1Y = 100

        Case 25
            GP1Y = 103

        Case 26
            GP1Y = 102

    End Select

End Function

Private Sub Form_Load()

    Dim nRet As Long
    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion(1, 1, 0, 0), True)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&

End Sub

Private Sub Form_Unload(Cancel As Integer)

    DeleteObject ResultRegion

End Sub

Private Sub Label1_Click(Index As Integer)

    Select Case Index

        Case 0
            'Play
            Call playmedia

        Case 1
            MMControl1.Command = "Pause"

        Case 2
            CommonDialog1.ShowOpen
            strfilename = CommonDialog1.FileName

        Case 3
            MMControl1.Command = "Stop"

    End Select

End Sub

Private Sub playmedia()

    MMControl1.Notify = False
    MMControl1.Shareable = False
    MMControl1.Wait = True
    MMControl1.FileName = strfilename
    MMControl1.Command = "Open"
    MMControl1.Command = "Prev"
    MMControl1.Command = "Play"
    Label2 = strfilename
    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

    If Label2.Left = 0 - Label2.Width Then

        Label2.Left = 6250

    Else

        Label2.Left = Label2.Left - 100

    End If

End Sub

