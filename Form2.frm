VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuDir 
         Caption         =   "&Add Directory"
      End
      Begin VB.Menu mnuSong 
         Caption         =   "Add &Song"
      End
      Begin VB.Menu asd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuDir_Click()
Form1.Drive1.Visible = True
Form1.Dir1.Visible = True
Form1.Command10.Visible = True
Form1.List1.Visible = False
Form1.List1.Text = ""
End Sub

Private Sub mnuExit_Click()
Unload Form1
Unload Form2

End Sub

Private Sub mnuSong_Click()
Form1.CommonDialog1.Filter = "MP3 File (*.mp3)|*.mp3|"
Form1.CommonDialog1.DialogTitle = "Open MP3 Song..."
Form1.CommonDialog1.ShowOpen

If Form1.CommonDialog1.CancelError = True Then
Form1.mp3.FileName = ""
End If


Form1.mp3.FileName = Form1.CommonDialog1.FileTitle
Form1.Text1.Text = Form1.CommonDialog1.FileTitle
Form1.List1.AddItem Form1.CommonDialog1.FileTitle
Form1.Command2.Enabled = True
Form1.Command6.Enabled = True
Form1.Caption = Form1.CommonDialog1.FileTitle
Form1.Text1.Text = Form1.List1.Text

If Form1.mp3.PlayState = mpPlaying Then
Form1.Timer1.Enabled = True
Form1.Label4.Caption = "File Size: " & FileLen(Form1.mp3.FileName) \ 1000 & " KB's"
Else
Form1.Timer1.Enabled = False
End If
Form1.mp3.Stop
End Sub


