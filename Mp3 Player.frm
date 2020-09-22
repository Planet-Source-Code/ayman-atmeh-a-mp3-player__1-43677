VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "MP3 Player"
   ClientHeight    =   8250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   Icon            =   "Mp3 Player.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   480
      TabIndex        =   30
      Top             =   6000
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1080
      TabIndex        =   29
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Mute Removed"
      Height          =   255
      Left            =   5040
      TabIndex        =   28
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Mute"
      Height          =   255
      Left            =   5040
      TabIndex        =   27
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   600
      TabIndex        =   26
      Top             =   1800
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   10
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture3 
      Height          =   735
      Left            =   5040
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   5280
      Top             =   5880
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      TickStyle       =   3
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   480
      Max             =   5000
      Min             =   -5000
      TabIndex        =   13
      Top             =   4800
      Width           =   2775
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   480
      Max             =   2500
      TabIndex        =   4
      Top             =   3960
      Value           =   2500
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Open"
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Center"
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   2760
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   5500
      Left            =   5280
      Top             =   4320
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3480
      Picture         =   "Mp3 Player.frx":0442
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   10
      ToolTipText     =   "Mute"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   6.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   9
      ToolTipText     =   "Stop"
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF0000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   5280
      Top             =   3120
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Play"
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3480
      Picture         =   "Mp3 Player.frx":088C
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   11
      ToolTipText     =   "RemoveMute"
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Play Controls"
      Height          =   735
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sound Countrols"
      Height          =   1575
      Left            =   240
      TabIndex        =   20
      Top             =   3600
      Width           =   4215
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Mute"
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Remove Mute"
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Play Time"
      Height          =   975
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   3615
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C0C0&
      Height          =   1035
      Left            =   840
      TabIndex        =   34
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add Songs"
      Height          =   2295
      Left            =   240
      TabIndex        =   31
      Top             =   5400
      Width           =   3735
      Begin VB.FileListBox File1 
         Height          =   870
         Left            =   720
         Pattern         =   "*.mp3"
         TabIndex        =   33
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   32
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1275
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "This Program Made By Ayman Atmeh"
      Top             =   7800
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Time"
      Height          =   615
      Left            =   3960
      TabIndex        =   15
      Top             =   1560
      Width           =   2295
      Begin VB.TextBox txttime 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   35
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   840
      Width           =   2775
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF0000&
      X1              =   120
      X2              =   6480
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF0000&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   8040
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      X1              =   6480
      X2              =   120
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      X1              =   6480
      X2              =   6480
      Y1              =   120
      Y2              =   8040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mp3 Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   960
      X2              =   4800
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   960
      X2              =   960
      Y1              =   720
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   4800
      X2              =   4800
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   960
      X2              =   4800
      Y1              =   720
      Y2              =   720
   End
   Begin MediaPlayerCtl.MediaPlayer mp3 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
On Error GoTo f
CommonDialog1.Filter = "mp3 Songs (*.mp3)|*.mp3|"
CommonDialog1.ShowOpen
CommonDialog1.DialogTitle = "Open MP3 Song"
mp3.FileName = CommonDialog1.FileName
If mp3.PlayState = mpPlaying Then
Timer1.Enabled = True
Label4.Caption = "File Size: " & FileLen(Form1.mp3.FileName) \ 1000 & " KB's"
Label5.Caption = "File Size: " & FileLen(mp3.FileName) \ 1000000 & " mB's"
Else
Timer1.Enabled = False
End If
mp3.FileName = CommonDialog1.FileTitle
Text1.Text = CommonDialog1.FileTitle
List1.AddItem CommonDialog1.FileTitle
Command2.Enabled = True
Command6.Enabled = True
Form1.Caption = CommonDialog1.FileTitle
Text1.Text = List1.Text & " " & Label3.Caption

If mp3.PlayState = mpPlaying Then
Timer1.Enabled = True
Label4.Caption = "File Size: " & FileLen(Form1.mp3.FileName) \ 1000 & " KB's"
Label5.Caption = "File Size: " & FileLen(mp3.FileName) \ 10000000 & " mB's"
Else
Timer1.Enabled = False
End If
Label5.Caption = "File Size: " & FileLen(mp3.FileName) \ 1000000 & " mB's"
mp3.Stop
f:
End Sub




Private Sub Command10_Click()
On Error GoTo f
Dim tel As Integer
File1.Path = Dir1.Path
If File1.ListCount <> 0 Then
    For tel = 1 To File1.ListCount
        File1.ListIndex = tel - 1
        
        If Len(Dir1.Path) > 3 Then
            List1.AddItem File1.FileName
        Else
           'Exit For
            'MsgBox "You can't add a drive, only folders", vbOKOnly, "Error"
           'Exit Sub
        List1.AddItem File1.FileName
        End If
    Next tel
Else
    MsgBox "No files were found in This folder", vbOKOnly, "Error"
End If

Drive1.Visible = False
Dir1.Visible = False
Command10.Visible = False
List1.Visible = True
f:
End Sub

Private Sub Command2_Click()
On Error GoTo f
mp3.Play
Timer1.Enabled = True
f:
End Sub







Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Form1.WindowState = 1
End Sub

Private Sub Command5_Click()
On Error GoTo f
mp3.Stop
f:
End Sub

Private Sub Command6_Click()
On Error GoTo f
mp3.Pause
f:
End Sub

Private Sub Command7_Click()
HScroll2.Value = 0
End Sub

Private Sub Command8_Click()
Text1.Text = "Mute"
End Sub

Private Sub Command9_Click()
Text1.Text = "Mute Removed"
End Sub


Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
If Button = 2 Then
PopupMenu Form2.mnufile
End If

End Sub

Private Sub Form_Load()
Label3.Visible = False
Text1.Text = "Welcome"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu Form2.mnufile
End If
End Sub

Private Sub Frame5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu Form2.mnufile
End If

End Sub

Private Sub HScroll1_Change()
Text1.Visible = True
Text1.Text = "0"
Dim pim, sha
sha = HScroll1.Value - 2500
mp3.Volume = sha
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = HScroll1.min
foo = HScroll1.Value
Text1.Text = "Volume " & foo \ 25 & " %"
hell:
Exit Sub

End Sub


Private Sub HScroll1_Scroll()
Text1.Visible = True
Dim pim, sha
sha = HScroll1.Value - 2500
mp3.Volume = sha
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = HScroll1.min
foo = HScroll1.Value
Text1.Text = "Volume " & foo \ 25 & " %"
hell:
Exit Sub

End Sub

Private Sub imgPlay_Click()
Command2_Click
Timer1.Enabled = True
End Sub

Private Sub HScroll2_Change()
On Error GoTo hell
'Text1.Visible = True
Text1.Visible = True
If HScroll2.Value > -500 And HScroll2.Value < 500 Then
Text1.Text = "Center"
End If
If HScroll2.Value < -500 Then
Text1.Text = "Balance:" & -(mp3.Balance / 50) & " % Left "
End If
If HScroll2.Value > 500 Then
Text1.Text = 0
Text1.Text = "Balance :" & mp3.Balance / 50 & " % Right "
End If
mp3.Balance = HScroll2.Value
hell:
Exit Sub
End Sub

Private Sub HScroll2_Scroll()
Text1.Visible = True

If HScroll2.Value > -2500 And HScroll2.Value < 2500 Then
Text1.Text = "Center"
End If
If HScroll2.Value < -2500 Then
Text1.Text = "Balance:" & -(mp3.Balance / 50) & " % Left "
End If
If HScroll2.Value > 2500 Then
Text1.Text = "Balance:" & mp3.Balance / 50 & " % Right "
End If
mp3.Balance = HScroll2.Value

End Sub



Private Sub List1_Click()
On Error GoTo f
mp3.FileName = List1.Text
mp3.Play
Form1.Caption = List1.Text
Label4.Caption = "File Size: " & FileLen(mp3.FileName) \ 1000 & " KB's"
Label5.Caption = "File Size: " & FileLen(mp3.FileName) \ 1000000 & " mB's"
Text1.Text = List1.Text & " " & txttime.Text
Timer1.Enabled = True
mp3.Play
mp3.AutoStart = True
f:
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo f
If Button = 2 Then
PopupMenu Form2.mnufile
End If
f:
End Sub

Private Sub Picture1_Click()
mp3.Mute = True
Picture2.Visible = True
Picture1.Visible = False
Label2.Visible = False
Label3.Visible = True
Text1.Visible = True
Command8_Click
End Sub

Private Sub Picture2_Click()
mp3.Mute = False
Picture2.Visible = False
Picture1.Visible = True
Label2.Visible = True
Label3.Visible = False
Text1.Visible = True
Command9_Click
End Sub



Private Sub Slider1_Scroll()
mp3.CurrentPosition = Slider1.Value
End Sub



Private Sub TabStrip1_Click()

End Sub

Private Sub Slider2_Change()
Text1.Visible = True
Text1.Text = "0"
Dim pim, sha
sha = Slider1.Value - 2500
mp3.Volume = sha
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = Slider1.min
foo = Slider1.Value
Text1.Text = "Volume " & foo \ 25 & " %"
hell:
Exit Sub

End Sub

Private Sub Slider2_Scroll()
Text1.Visible = True
Dim pim, sha
sha = Slider1.Value - 2500
mp3.Volume = sha
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = Slider1.min
foo = Slider1.Value
Text1.Text = "Volume " & foo \ 25 & " %"
hell:
Exit Sub
End Sub

Private Sub Timer1_Timer()
On Error GoTo f
Slider1.Value = mp3.CurrentPosition
Slider1.Max = mp3.Duration
ProgressBar1.Value = mp3.CurrentPosition
ProgressBar1.Max = mp3.Duration
If mp3.Duration > 0 Then
Else
Exit Sub
End If
Dim i As Integer
Dim min As Integer
Dim sec As Integer
i = Val(Format(mp3.CurrentPosition, "###"))
If i > 59 Then
min = i \ 60
sec = i Mod 60
txttime.Text = Format(min, "0#") & ":" & Format(sec, "00")
Else
If i > -1 Then
txttime.Text = "00" & ":" & Format(i, "0#")
End If
End If

i = Val(Format(Form1.mp3.Duration, "###"))
If i > 59 Then
min = i \ 60
sec = i Mod 60
Text2.Text = "/ " & " " & Format(min, "0#") & ":" & Format(sec, "00")
Else
If i > -1 Then
Label3.Caption = "/ " & " 00" & ":" & Format(i, "0#")
End If
End If
f:
End Sub


Private Sub Timer2_Timer()
Text1.Visible = False
End Sub

Private Sub Timer3_Timer()
If Text3.Left < Picture1.Width - Picture1.Width - Text3.Width Then
Text3.Left = Picture1.Width - 1
Text3.Left = Text3.Left - 5
Else
   Text3.Left = Text3.Left - 10
End If

End Sub


