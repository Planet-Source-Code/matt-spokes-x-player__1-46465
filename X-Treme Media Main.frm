VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Player 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "X-Player By SpOkEsY"
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   Icon            =   "X-Treme Media Main.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "X-Treme Media Main.frx":0CCA
   ScaleHeight     =   5520
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Mute"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   1920
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Max             =   2500
      SelStart        =   2500
      TickStyle       =   3
      Value           =   2500
   End
   Begin VB.OptionButton Normal 
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton Shuffle 
      Caption         =   "Shuffle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   1695
   End
   Begin VB.OptionButton Repeat 
      Caption         =   "Repeat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save Playlist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton Load 
      Caption         =   "Load Playlist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton Clear 
      Caption         =   "Clear Playlist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Del 
      Caption         =   "Remove Song"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add Song"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5640
      Top             =   5640
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5640
      Top             =   5640
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1200
      ItemData        =   "X-Treme Media Main.frx":6D792
      Left            =   120
      List            =   "X-Treme Media Main.frx":6D794
      TabIndex        =   5
      Top             =   4080
      Width           =   5535
   End
   Begin VB.CommandButton Stopper 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Pause 
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Play 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Slide 
      Height          =   240
      Left            =   120
      Picture         =   "X-Treme Media Main.frx":6D796
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   5520
      X2              =   240
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label Close 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   4800
      TabIndex        =   19
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Min 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   3960
      TabIndex        =   18
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   120
      Picture         =   "X-Treme Media Main.frx":6DC80
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "100 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   2040
      Width           =   615
   End
   Begin VB.Shape SliderBar 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current Status : Stopped"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00 / 00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   5760
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000F&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   5535
   End
End
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tx As Integer, Ty As Integer, DN As Boolean  'this is for the seekbar
Dim Txa As Integer, DNa As Boolean
Dim Tyb, DNb As Boolean
Function ConvertTime(i As Integer) ' this is the time function
Secs = i Mod 60
Mins = Int(i / 60) Mod 60
If Secs < 10 Then Secs = "0" & Secs
If Mins < 10 Then Mins = "0" & Mins
ConvertTime = Mins & ":" & Secs
End Function

Private Sub Add_Click() 'this is the add file button
On Error Resume Next
CommonDialog1.Filter = "Audio Files|*.wav;*.mid;*.mp3;mp2;*.mod;*.wma;|"
CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Add File"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen

List1.AddItem CommonDialog1.FileTitle
Del.Enabled = True
Play.Enabled = True
End Sub

Private Sub Check1_Click() 'this is mute
If MediaPlayer1.Mute = True Then
MediaPlayer1.Mute = False
Else
MediaPlayer1.Mute = True
End If
End Sub



Private Sub Clear_Click() 'this is clear
ask = MsgBox("Do you want to clear your list ?", vbQuestion + vbYesNo, "Clear Playlist")
If ask = vbYes Then
MediaPlayer1.Stop
List1.Clear
Else
End If
End Sub

Private Sub Close_Click() 'this is exit
End
End Sub

Private Sub Del_Click() 'this is to delete
If List1.ListIndex > -1 Then
On Error Resume Next
If List1.Text = MediaPlayer1.FileName Then MsgBox "You can't remove the file you're playing": Exit Sub
List1.RemoveItem List1.ListIndex
End If


End Sub

Private Sub Form_Load() 'this is various things
Timer1.Interval = 1000
Timer2.Interval = 1000
Slider2.Max = 2500
Stopper.Enabled = False
Play.Enabled = False
Pause.Enabled = False
Del.Enabled = False
End Sub



Private Sub List1_DblClick() 'this for dbl clicking the list item
Play.Value = True
End Sub

Private Sub Load_Click() 'this is for playlists
On Error Resume Next
On Error GoTo err
Close #1
Dim X
OpenFile:
CommonDialog1.Filter = "All Supported|*.m3u;*.pls|"
CommonDialog1.DialogTitle = "Open List"
CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.ShowOpen
CommonDialog1.CancelError = True
Open CommonDialog1.FileName For Input As #1
List1.Clear
Do
Input #1, X
List1.AddItem (X)
Loop

Close #1
err:
Exit Sub
End Sub
Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long) 'this is when the song ends

If Repeat.Value = True Then
MediaPlayer1.Play
Else
End If
If Shuffle.Value = True Then
If List1.ListCount = 0 Then
 MsgBox "There are no files to play!", vbCritical
 Exit Sub
 End If
 Randomize Timer
 myvalue = Int((List1.ListCount * Rnd))
    List1.ListIndex = myvalue
    MediaPlayer1.FileName = List1.Text
    Text1.Caption = List1.Text
    If List1.Text <> "" Then
        MediaPlayer1.Play
        Text1.Caption = List1.Text
        Exit Sub
    End If



End If


If Normal.Value = True Then
If List1.ListIndex = List1.ListCount - 1 Then List1.ListIndex = -1
    List1.ListIndex = List1.ListIndex + 1



 End If

        MediaPlayer1.FileName = List1.Text
        MediaPlayer1.Play
Text1.Caption = List1.Text
    
        If MediaPlayer1.PlayState = mpStopped Then MediaPlayer1.Play




End Sub

Private Sub MediaPlayer1_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long) 'this is for the status label
    Dim curState As String
    Select Case NewState
        Case 0
            curState = "Stopped"
        Case 1
            curState = "Paused"
        Case 2
            curState = "Playing"
    End Select
    Label2.Caption = "Current Status : " & curState
End Sub



Private Sub Min_Click() 'this is min to sys tray
Me.WindowState = 1
End Sub

Private Sub Pause_Click() 'this is pause
On Error Resume Next
If MediaPlayer1.PlayState = mpPlaying Then
MediaPlayer1.Pause
Play.Enabled = False
Stopper.Enabled = False
Else
MediaPlayer1.Play
Play.Enabled = True
Stopper.Enabled = True
End If
End Sub

Private Sub Play_Click() 'this is play
On Error Resume Next
MediaPlayer1.FileName = List1.Text
MediaPlayer1.Play
Text1.Caption = MediaPlayer1.FileName
Play.Enabled = True
Stopper.Enabled = True
Pause.Enabled = True
End Sub



Private Sub Save_Click() 'save playlist
On Error Resume Next
CommonDialog1.Filter = "Playlist File (M3u)|*.m3u|Playlist File (Pls)|*.pls"
CommonDialog1.DialogTitle = "Save List"
CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.ShowSave
CommonDialog1.CancelError = True

Open CommonDialog1.FileName For Output As #1
For X = 0 To List1.ListCount - 1
Print #1, List1.List(X)
Next X

Close #1
End Sub



Private Sub SliderBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'for   clicking the sliderbar
If Button = 1 Then

MediaPlayer1.CurrentPosition = MediaPlayer1.Duration / SliderBar.Width
End If
End Sub
Private Sub Slide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) ' for slidin the slider
    DNa = True
    Txa = X


End Sub
Private Sub Slide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'also for slidin the slider
    If DNa Then
        NewLeft = Slide.Left + X - Txa
        If NewLeft < SliderBar.Left + 3 Then
            NewLeft = SliderBar.Left + 3
        End If
        If NewLeft > SliderBar.Width + SliderBar.Left - 7 - Slide.Width Then
            NewLeft = SliderBar.Width + SliderBar.Left - 7 - Slide.Width
        End If
        Slide.Left = NewLeft
    End If
End Sub
Private Sub Slide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'and again  for the slider
    Dim offseti As Single
    DNa = False
    offseti = (Slide.Left - SliderBar.Left - 3) / (SliderBar.Width - 10 - Slide.Width)
    MediaPlayer1.CurrentPosition = Int(MediaPlayer1.Duration * offseti)
End Sub




Private Sub Slider2_Scroll() 'volume
Dim a As Integer, b As Integer
Dim d, c
c = Slider2.Value - 2500
MediaPlayer1.Volume = c
b = Slider2.Min
a = Slider2.Value
Label3.Caption = a \ 25 & " %"
End Sub

Private Sub Stopper_Click() 'stop!
On Error Resume Next
MediaPlayer1.CurrentPosition = 0
MediaPlayer1.Stop
Pause.Enabled = False
Stopper.Enabled = False

End Sub

Private Sub Timer1_Timer() 'another time function
If MediaPlayer1.PlayState = mpPlaying Then
Label1.Caption = ConvertTime(Round(MediaPlayer1.CurrentPosition, 0)) & " / " & ConvertTime(Round(MediaPlayer1.Duration, 0))
Else
Label1.Caption = "00:00 / 00:00"
End If
End Sub

Private Sub Timer2_Timer() 'for the slider
Dim tm As Integer, tt As Integer, tp As Single, offset As Integer
    tm = Int(MediaPlayer1.CurrentPosition)
    tt = Int(MediaPlayer1.Duration)
    If tm <> -1 Then
        tp = tm / tt
        offset = Int((SliderBar.Width - 10 - Slide.Width) * tp)
        If Not DNa Then Slide.Left = offset + SliderBar.Left + 3
    Else
    End If
End Sub
