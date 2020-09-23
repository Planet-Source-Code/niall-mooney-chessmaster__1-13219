VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ni-Star Enterprises - ChessMASTER Server Mode"
   ClientHeight    =   7620
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8325
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ChessM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   555
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   12
      Left            =   2085
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   38
      Top             =   4515
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Accept Popup Messages"
      Height          =   255
      Left            =   5880
      TabIndex        =   36
      Top             =   6000
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Popup"
      Height          =   255
      Left            =   6960
      TabIndex        =   35
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Board"
      Height          =   1215
      Left            =   5880
      TabIndex        =   31
      Top             =   6360
      Width           =   2415
      Begin VB.OptionButton Option1 
         Caption         =   "Alternative Board"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   34
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Original but better"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Original and best"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.TextBox Terminal 
      Height          =   2835
      Left            =   5835
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   2145
      Width           =   2430
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6120
      Top             =   6480
   End
   Begin VB.TextBox Handle 
      Height          =   285
      Left            =   5820
      TabIndex        =   24
      Top             =   1785
      Width           =   2445
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1920
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".Nif"
      DialogTitle     =   "NiFile systems for ChessMaster"
      FileName        =   "*.Nif"
      Filter          =   "*.Nif"
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   11
      Left            =   2100
      Picture         =   "ChessM.frx":030A
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   23
      Top             =   4530
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   10
      Left            =   2100
      Picture         =   "ChessM.frx":0798
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   22
      Top             =   4530
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   9
      Left            =   2100
      Picture         =   "ChessM.frx":0C59
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   21
      Top             =   4530
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   8
      Left            =   2100
      Picture         =   "ChessM.frx":11DE
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   20
      Top             =   4530
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   7
      Left            =   2100
      Picture         =   "ChessM.frx":16A4
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   19
      Top             =   4530
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   6
      Left            =   2100
      Picture         =   "ChessM.frx":1B82
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   18
      Top             =   4530
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   5
      Left            =   2100
      Picture         =   "ChessM.frx":2048
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   17
      Top             =   4530
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   4
      Left            =   2100
      Picture         =   "ChessM.frx":24D6
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   16
      Top             =   4530
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   3
      Left            =   2100
      Picture         =   "ChessM.frx":2997
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   15
      Top             =   4530
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   2
      Left            =   2100
      Picture         =   "ChessM.frx":2F1C
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   14
      Top             =   4530
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   1
      Left            =   2100
      Picture         =   "ChessM.frx":33E2
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   13
      Top             =   4530
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   0
      Left            =   2100
      Picture         =   "ChessM.frx":38C0
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   12
      Top             =   4530
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Send"
      Height          =   240
      Left            =   5835
      TabIndex        =   9
      Top             =   5655
      Width           =   1035
   End
   Begin VB.TextBox Chattxt 
      Height          =   585
      Left            =   5835
      TabIndex        =   8
      Top             =   5025
      Width           =   2430
   End
   Begin MSWinsockLib.Winsock Ws 
      Left            =   1920
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   500
      LocalPort       =   499
   End
   Begin VB.Frame Frame1 
      Caption         =   "WinSocket Control"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton Command2 
         Caption         =   "Resign"
         Height          =   255
         Left            =   4680
         TabIndex        =   37
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Disconnect"
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Refresh Board"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Send..."
         Height          =   255
         Left            =   3120
         TabIndex        =   27
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Chat..."
         Height          =   255
         Left            =   4320
         TabIndex        =   26
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Offline Mode"
         Height          =   255
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Load..."
         Height          =   255
         Left            =   4320
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save As"
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Oiplbl 
         Caption         =   "Opponent I.P."
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label oNmlbl 
         Caption         =   "Opponent Name"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Nmlbl 
         Caption         =   "Your Name"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label IPlbl 
         Caption         =   "Your I.P."
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5790
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   8
      ScaleMode       =   0  'User
      ScaleWidth      =   8
      TabIndex        =   10
      Top             =   1830
      Width           =   5790
   End
   Begin VB.Label Label1 
      Caption         =   "Your Handle is:"
      Height          =   255
      Left            =   5820
      TabIndex        =   25
      Top             =   1545
      Width           =   1695
   End
   Begin VB.Label Coords 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      TabIndex        =   11
      Top             =   105
      Width           =   2400
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuNew 
         Caption         =   "&New Offline Session"
      End
      Begin VB.Menu MnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu MnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDisconnect 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu MnuSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSave 
         Caption         =   "&Save As"
      End
      Begin VB.Menu MnuResign 
         Caption         =   "&Resign"
      End
      Begin VB.Menu MnuSeperator4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuFeatures 
      Caption         =   "&Features"
      Begin VB.Menu MnuSend 
         Caption         =   "&Send File to Client"
      End
      Begin VB.Menu MnuChat 
         Caption         =   "&Chat on ChessMASTER WebSite"
      End
      Begin VB.Menu MnuTrans 
         Caption         =   "Translucency Colour"
         Begin VB.Menu MnuBlue 
            Caption         =   "Blue"
         End
         Begin VB.Menu MnuCyan 
            Caption         =   "Cyan"
         End
         Begin VB.Menu MnuGreen 
            Caption         =   "Green"
         End
         Begin VB.Menu MnuMagenta 
            Caption         =   "Magenta"
         End
         Begin VB.Menu MnuRed 
            Caption         =   "Red"
         End
         Begin VB.Menu MnuYellow 
            Caption         =   "Yellow"
         End
         Begin VB.Menu MnuWhite 
            Caption         =   "Transparent"
         End
      End
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "&About"
      NegotiatePosition=   3  'Right
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Board(8, 8) As String
Dim Mode As String, Selected%, SelX%, SelY%, ToX%, ToY%
Dim Turn As String, TransColour  As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Sub UpdateConnectionStats()
IPlbl.Caption = "Your IP: " & Ws.LocalIP
Nmlbl.Caption = "Your Name: " & Ws.LocalHostName
Oiplbl.Caption = "Opponent IP: " & Ws.RemoteHostIP
oNmlbl.Caption = "Opponent Name: " & Ws.RemoteHost
Handle.Text = Ws.LocalHostName
If Mode = "Offline" Then 'Show that you're playing ChessMASTER
Oiplbl.Caption = "Opponent IP: " & Ws.LocalIP
oNmlbl.Caption = "Opponent Name: ChessMASTER"
End If
If Mode = "" Then 'Could be offline two player
Oiplbl.Caption = "Opponent IP: " & Ws.LocalIP
oNmlbl.Caption = "Opponent Name: Player 2"
End If
End Sub

Private Sub Chattxt_Change()
Command7.Default = True
End Sub

Private Sub Command1_Click()
If Mode <> "Online" Then Exit Sub

On Error Resume Next
If Chattxt.Text <> "" Then
Ws.SendData "popp> " & Ws.LocalHostName & ">> " & Chattxt.Text
Terminal.Text = Handle.Text & ">> " & Chattxt.Text & vbNewLine & Terminal.Text
Chattxt.Text = ""
End If
If Err Then MsgBox Err.Description, vbCritical, "NSE ChessMaster"
Command7.Default = False
End Sub

Private Sub Command10_Click()
If Mode <> "Online" Then Exit Sub

On Error Resume Next
CD.DefaultExt = ".*"
CD.Filter = "*.*"
CD.FileName = "*.*"
CD.ShowOpen
i = MsgBox("Send File '" + CD.FileName + "' (" + Str(FileLen(CD.FileName)) + "Bytes) to " + Ws.RemoteHost, vbQuestion + vbYesNo, "NSE ChessMASTER")
If i = vbYes Then
Dim WholeFile$
'WholeFile$ = Space(FileLen(CD.FileName))
Ln$ = Space(FileLen(CD.FileName))
Open CD.FileName For Input As #7
Do
Input #7, Ln$
WholeFile$ = WholeFile$ + Ln$
Loop Until EOF(7)
Ws.SendData WholeFile$
Close #7
End If
If Err Then MsgBox Err.Description, vbCritical, "NSE ChessMASTER"
CD.DefaultExt = ".Nif"
CD.Filter = "*.Nif"
End Sub

Private Sub Command2_Click()
If Mode = "" Then Exit Sub

i = MsgBox("Are you sure you want to resign", vbQuestion + vbYesNo, "NSE ChessMASTER")
If i = vbYes Then
If Mode = "Online" Then Ws.SendData "Vctry"
Mode = ""
For x = 0 To 8
For y = 0 To 8
Board(x, y) = ""
Next
Next
Picture1.Cls
MsgBox "You lose", vbExclamation, "NSE ChessMASTER"
End If
End Sub

Private Sub Command3_Click()
If Mode <> "Online" Then Exit Sub
Ws.SendData "Vctry"
MsgBox "You lose", vbExclamation, "NSE ChessMASTER"
Ws.Close
End Sub

Private Sub Command4_Click()
On Error Resume Next
CD.ShowSave
Open CD.FileName For Output As #5
Print #5, Turn
For x = 0 To 8
For y = 0 To 8
Print #5, Board(x, y)
Next
Next
Close #5
If Err Then MsgBox Err.Description, vbCritical, "NSE ChessMaster"
DrawBoard
End Sub

Private Sub Command5_Click()

CD.FileName = "*.Nif"
CD.ShowOpen

If UCase$(Right$(CD.FileName, 4)) <> ".NIF" Then Exit Sub

If Mode = "Offline" Then 'Don't load while playing AI
MsgBox "Please resign before loading a new board", vbExclamation, "NSE ChessMASTER Ni-File Algorithmn"
Exit Sub
End If

If Mode = "" Then 'If there is no game in progress
i = MsgBox("Begining Practice session with loaded game." & vbNewLine & "Do you wish to play ChessMASTER on this board?" & vbNewLine & "If not you may play in Offline two player mode", vbYesNoCancel + vbQuestion, "NSE ChessMASTER Ni-File Algorithmn")
If i = vbYes Then 'The user wants to play the AI on the board
Mode = "Offline"
Turn = "W"
Timer1.Enabled = True
ElseIf i = vbCancel Then 'The user canceled loading
Exit Sub
End If
End If

Open CD.FileName For Input As #5

If LOF(5#) <> 351 Then 'Wrong length of FILE !
MsgBox "Error: Incorrect File Format. This loading algotithmn accepts only Spec2CM Ni-Files", vbCritical + vbSystemModal, "NSE ChessMASTER Ni-File Algorithmn"
Close #5
Exit Sub
End If

Line Input #5, Turn
Trim$ (Turn)
For x = 0 To 8
For y = 0 To 8
Line Input #5, j$
Board(x, y) = Trim(j$)
Next
Next
Close #5
DrawBoard
'Send the Board to the other player (client)
If Mode = "Online" Then
SendString$ = "load>"
SendString$ = SendString$ + Trim$(Turn)
For x = 0 To 8
For y = 0 To 8
NewBit$ = Trim$(Board(x, y)) + Space$(7 - Len(Board(x, y)))
SendString$ = SendString$ + NewBit$
Next
Next
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
If Mode <> "" Then
If MsgBox("Start a new Offline Mode session", vbMsgBoxSetForeground + vbQuestion + vbYesNo, "NSE ChessMASTER") = vbNo Then Exit Sub
End If
initBoard
DrawBoard
Mode = "Offline"
Turn = "W"
Timer1.Enabled = True
End Sub

Private Sub Command7_Click()
If Mode <> "Online" Then Exit Sub

On Error Resume Next
If Chattxt.Text <> "" Then
Ws.SendData "chat> " & Ws.LocalHostName & ">> " & Chattxt.Text
Terminal.Text = Handle.Text & ">> " & Chattxt.Text & vbNewLine & Terminal.Text
Chattxt.Text = ""
End If
If Err Then MsgBox Err.Description, vbCritical, "NSE ChessMaster"
Command7.Default = False
End Sub

Private Sub Command8_Click()
DrawBoard
End Sub

Private Sub Command9_Click()
Shell "explorer.exe Http://www.homestead.com/nmooney/chatpage.html", vbMinimizedFocus
End Sub

Private Sub Form_Load()
Randomize Timer

CD.InitDir = App.Path
Me.Show

Ws.Listen
Command6.SetFocus

TransColour = vbWhite
Option1_Click 1
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
UpdateConnectionStats
End Sub

Private Sub MnuAbout_Click()
frmAbout.Show
End Sub

Private Sub MnuBlack_Click()
TransColour = vbBlack
DrawBoard
End Sub

Private Sub MnuBlue_Click()
TransColour = vbBlue
DrawBoard
End Sub

Private Sub MnuChat_Click()
Command9_Click
End Sub

Private Sub MnuCyan_Click()
TransColour = vbCyan
DrawBoard
End Sub

Private Sub MnuDisconnect_Click()
Command3_Click
End Sub

Private Sub MnuExit_Click()
If Mode <> "" Then
MsgBox "Please end any running game session before leaving", vbInformation, "NSE ChessMASTER"
Else
Unload Me
End
End If
End Sub

Private Sub MnuGreen_Click()
TransColour = vbGreen
DrawBoard
End Sub

Private Sub MnuMagenta_Click()
TransColour = vbMagenta
DrawBoard
End Sub

Private Sub MnuNew_Click()
Command6_Click
End Sub

Private Sub MnuOpen_Click()
Command5_Click
End Sub

Private Sub MnuRed_Click()
TransColour = vbRed
DrawBoard
End Sub

Private Sub MnuResign_Click()
Command2_Click
End Sub

Private Sub MnuSave_Click()
Command4_Click
End Sub

Private Sub MnuSend_Click()
Command10_Click
End Sub

Private Sub MnuWhite_Click()
TransColour = vbWhite
DrawBoard
End Sub

Private Sub MnuYellow_Click()
TransColour = vbYellow
DrawBoard
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
Picture1.Picture = LoadResPicture(101, 0)
Case 1
Picture1.Picture = LoadResPicture(102, 0)
Case 2
Picture1.Picture = LoadResPicture(103, 0)
End Select
DrawBoard
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Selected% = 0 Then   'Selecting a piece to move
If Turn = Right(Board(Int(x + 1), Int(y + 1)), 1) And Board(Int(x + 1), Int(y + 1)) <> "" Then 'If your turn
Selected% = 1
SelX% = Int(x + 1)
SelY% = Int(y + 1)
DrawBoard
If Mode = "Online" Then
If Right(Board(SelX%, SelY%), 1) = "W" Then Selected% = 0
End If
Exit Sub
End If
End If

If Selected% = 1 Then  'Selecting a place to move to
If ToX% = SelX% And ToY% = SelY% Then Selected% = 0
DrawBoard
ToX% = Int(x + 1)
ToY% = Int(y + 1)
MoveIfAlright
End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

ToX% = Int(x + 1)
ToY% = Int(y + 1)

If Int(x + 1) > 8 Or Int(y + 1) > 8 Then Exit Sub
Coords.Caption = "Coords: " & Int(x + 1) & ", " & Int(y + 1) & vbNewLine & "Piece: " & Board(Int(x + 1), Int(y + 1))

If Mode = "Offline" And Turn = "W" Then
Coords.Caption = Coords.Caption & vbNewLine & "Waiting For ChessMASTER"
Exit Sub
End If

If Selected% = 1 Then Coords.Caption = Coords.Caption & vbNewLine & Mid(Board(SelX%, SelY%), 1, Len(Board(SelX%, SelY%)) - 1) & " @ " & SelX% & ", " & SelY% & " to " & ToX% & ", " & ToY%
If Selected% = 1 And MoveAlright(SelX%, SelY%, ToX%, ToY%) Then Coords.Caption = Coords.Caption & vbNewLine & "-LEGAL MOVE-"

If Mode <> "" And Turn = "W" Then Coords.Caption = Coords.Caption & vbNewLine & "Waiting For " & Ws.RemoteHost
If Mode <> "" And Turn = "B" Then Coords.Caption = Coords.Caption & vbNewLine & "Your Turn, " & Ws.LocalHostName
If Mode = "" And Turn = "W" Then Coords.Caption = Coords.Caption & vbNewLine & "Waiting For Player 2"
If Mode = "" And Turn = "B" Then Coords.Caption = Coords.Caption & vbNewLine & "Your Turn, " & Ws.LocalHostName
End Sub

Private Sub Timer1_Timer()
If Mode <> "Offline" Then Exit Sub
If Turn = "B" Then Exit Sub
'Declarations of 5 parrellel arrays for AI operations
Dim HypFrX%(72)
Dim HypFrY%(72)
Dim HypToX%(72)
Dim HypToY%(72)
Dim HypQuality(72)

'AI. This AI uses the actual board memory & therfor cannot be recursive above 1 level.
'If it used a copy of the board it could minipulate and play ahead (recusivly).

'Scan board and move all possibilities giving a score for taking enemy peices
For x = 1 To 8
For y = 1 To 8
If Right(Board(x, y), 1) = "W" And Board(x, y) <> "" Then
For X2 = 1 To 8
For Y2 = 1 To 8
If MoveAlright(x, y, X2, Y2) Then
HypFrX%(X2 * 8 - 1 + Y2) = x
HypFrY%(X2 * 8 - 1 + Y2) = y
HypToX%(X2 * 8 - 1 + Y2) = X2
HypToY%(X2 * 8 - 1 + Y2) = Y2
HypQuality(X2 * 8 - 1 + Y2) = QualityOfTarget(X2, Y2)
End If
Next
Next
End If
Next
Next
'Scan board moving all possibities of enemies and subtracting score from move locations
For x = 1 To 8
For y = 0 To 8
If Right(Board(x, y), 1) = "W" Then
For X2 = 1 To 8
For Y2 = 1 To 8
If MoveAlright(x, y, X2, Y2) Then
HypQuality(X2 * 8 - 1 + Y2) = HypQuality(X2 * 8 - 1 + Y2) - CostOfLoss(x, y)
End If
Next
Next
End If
Next
Next

'Scan accumulated score & coordinate arrays & execute the one with the greatest score
NumberOfBest = 0
BestSoFar = -100
For i = 1 To 72
If Not HypQuality(i) = Empty And HypQuality(i) > BestSoFar Or (HypQuality(i) = BestSoFar And Rnd < 0.2) Then
NumberOfBest = i
BestSoFar = HypQuality(i)
End If
Next

SelX% = HypFrX%(NumberOfBest)
SelY% = HypFrY%(NumberOfBest)
ToX% = HypToX%(NumberOfBest)
ToY% = HypToY%(NumberOfBest)

MoveIfAlright
End Sub

Private Sub Ws_Connect()
Me.Caption = "Ni-Star Enterprises - ChessMASTER Client Mode CONNECTED"
UpdateConnectionStats
End Sub

Private Sub Ws_ConnectionRequest(ByVal requestID As Long)
i = MsgBox("Do you accept a connection request from " & requestID & "?", vbQuestion + vbYesNo, "NSE Chess Master")
If i = vbYes Then Ws.Accept requestID
If i = vbNo Then
i = MsgBox("Are you certain you wish to turn away" & requestID & "?", vbQuestion + vbYesNo, "NSE Chess Master")
If i = vbNo Then Ws.Accept requestID
End If
Mode = "Online"
InitiateOnlineMode
Turn = "B"
End Sub

Private Sub Ws_DataArrival(ByVal bytesTotal As Long)
Ws.GetData DataReceived
DataReceived = Str(DataReceived)

If Left$(DataRecieved, 5) = "Vctry" Then
MsgBox "You win. " & Ws.RemoteHost & "either resigned of lost the match!", vbExclamation, "NSE ChessMASTER"
End If

If Left$(DataReceived, 5) = "popp>" Then 'It's a popup message
If Check1.Value = 1 Then
MsgBox Mid$(DataReceived, 5), vbOKOnly + vbSystemModal + vbInformation, "NSE ChessMASTER Popup from " & Ws.RemoteHost
Else
Terminal.Text = Mid$(DataReceived, 5) & vbNewLine & "Popup Message: " & Terminal.Text
End If
End If

If Left$(DataReceived, 5) = "chat>" Then 'It's a chat message
Terminal.Text = Mid$(DataReceived, 5) & vbNewLine & Terminal.Text
End If

If Left$(DataReceived, 5) = "move>" Then 'It's a move command
SelX% = Mid(DataReceived, 5, 1)
SelY% = Mid(DataReceived, 6, 1)
ToX% = Mid(DataReceived, 7, 1)
ToY% = Mid(DataReceived, 8, 1)
MoveIfAlright
End If
End Sub

Sub initBoard()
For x = 0 To 8
For y = 0 To 8
Board(x, y) = ""
Next
Next
'Black Pieces
Board(1, 1) = "CastleW"
Board(2, 1) = "KnightW"
Board(3, 1) = "BishopW"
Board(4, 1) = "QueenW"
Board(5, 1) = "KingW"
Board(6, 1) = "BishopW"
Board(7, 1) = "KnightW"
Board(8, 1) = "CastleW"
For i = 1 To 8: Board(i, 2) = "PawnW": Next
'White Pieces
Board(1, 8) = "CastleB"
Board(2, 8) = "KnightB"
Board(3, 8) = "BishopB"
Board(4, 8) = "QueenB"
Board(5, 8) = "KingB"
Board(6, 8) = "BishopB"
Board(7, 8) = "KnightB"
Board(8, 8) = "CastleB"
For i = 1 To 8: Board(i, 7) = "PawnB": Next
End Sub

Sub DrawBoard()
Picture1.Cls
For x = 0 To 8
For y = 0 To 8
If Board(x, y) <> "" Then
n = 0
Select Case Mid(Board(x, y), 1, Len(Board(x, y)) - 1)
Case "King": n = 0
Case "Queen": n = 1
Case "Bishop": n = 2
Case "Knight": n = 3
Case "Castle": n = 4
Case "Pawn": n = 5
End Select
If Right(Board(x, y), 1) = "W" Then n = n + 6
'Picture1.PaintPicture Piece(n).Picture, x - 1.25, y - 1.25, , , , , , , vbSrcCopy
Piece(n).BackColor = Picture1.Point(x - 0.125, y - 0.125)
If Option1(0).Value = False Then Piece(n).BackColor = TransColour
If Selected% And x = SelX% And y = SelY% Then Piece(n).BackColor = vbGreen
If (TransColour = vbGreen Or TransColour = vbYellow) And Selected% And x = SelX% And y = SelY% Then Piece(n).BackColor = vbBlue
BitBlt Picture1.hDC, (x * 48) - 48, (y * 48) - 48, 48, 48, Piece(n).hDC, 0, 0, vbSrcAnd
If Picture1.Point(x - 0.125, y - 0.125) = vbBlack Then BitBlt Picture1.hDC, (x * 48) - 48, (y * 48) - 48, 48, 48, Piece(n).hDC, 0, 0, vbSrcPaint
Else 'No unit to draw
Piece(12).BackColor = TransColour
If Option1(0).Value = False Then BitBlt Picture1.hDC, (x * 48) - 48, (y * 48) - 48, 48, 48, Piece(12).hDC, 0, 0, vbSrcAnd
End If
Next
Next
End Sub

Sub InitiateOnlineMode()
initBoard
Timer1.Enabled = False
End Sub

Sub MoveIfAlright()
'On Error Resume Next
If Board(SelX%, SelY%) = "" Then Exit Sub
Colour$ = Right(Board(SelX%, SelY%), 1)
'See if were trying to take a piece
If Right(Board(SelX%, SelY%), 1) <> Right(Board(ToX%, ToY%), 1) And Right(Board(ToX%, ToY%), 1) <> "" Then taking = True
'See if were trying to take own piece!!
If Right(Board(SelX%, SelY%), 1) = Right(Board(ToX%, ToY%), 1) And Right(Board(ToX%, ToY%), 1) <> "" Then Exit Sub
'See if moving (to blank space)
If Board(ToX%, ToY%) = "" Then moving = True

Select Case Mid(Board(SelX%, SelY%), 1, Len(Board(SelX%, SelY%)) - 1)
Case "Pawn"     'Prawn movement
If moving And (Board(ToX%, ToY%) = "" And ToY% = SelY% - 1 And Colour$ = "B" And SelX% = ToX%) Then Ok = 1    'White vertical
If moving And (Board(ToX%, ToY%) = "" And ToY% = SelY% + 1 And Colour$ = "W" And SelX% = ToX%) Then Ok = 3    'Black Vertical
If taking And ToY% = SelY% - 1 And Colour$ = "B" And (ToX% = SelX% - 1 Or ToX% = SelX% + 1) Then Ok = 5  'Black Up + Left/Right to White Taking a piece
If taking And ToY% = SelY% + 1 And Colour$ = "W" And (ToX% = SelX% - 1 Or ToX% = SelX% + 1) Then Ok = 5  'White Up + Left/Right to Black Taking a piece
If moving And (Colour$ = "B" And SelY% = 7 And ToY% = SelY% - 2) And ToX% = SelX% Then Ok = 1  'First move may be double BLACK
If moving And (Colour$ = "W" And SelY% = 2 And ToY% = SelY% + 2) And ToX% = SelX% Then Ok = 3  'First move may be double BLACK
If Ok = 0 Then Exit Sub
Case "King"     'King movement
If (taking Or moving) And ToX% <= SelX% + 1 And ToX% >= SelX% - 1 And ToY% <= SelY% + 1 And ToY% >= SelY% - 1 Then Ok = 5 'Move in any direction by one
If Ok = 0 Then Exit Sub
Case "Queen"
'Diagonal
If (taking Or moving) And SelX% - ToX% = SelY% - ToY% Or SelX% - ToX% = ToY% - SelY% Then Ok = 5
'Left
If (taking Or moving) And ToX% = SelX% - 1 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 2 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 3 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 4 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 5 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 6 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 7 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 8 And ToY% = SelY% Then Ok = 4
'Right
If (taking Or moving) And ToX% = SelX% + 1 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 2 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 3 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 4 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 5 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 6 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 7 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 8 And ToY% = SelY% Then Ok = 2
'Up
If (taking Or moving) And ToY% = SelY% - 1 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 2 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 3 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 4 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 5 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 6 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 7 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 8 And ToX% = SelX% Then Ok = 1
'Down
If (taking Or moving) And ToY% = SelY% + 1 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 2 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 3 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 4 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 5 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 6 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 7 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 8 And ToX% = SelX% Then Ok = 3
If Ok = 0 Then Exit Sub
Case "Bishop"
If (taking Or moving) And (ToX% = SelX% + 1 And ToY% = SelY% + 1) Or (ToX% = SelX% - 1 And ToY% = SelY% - 1) Or (ToX% = SelX% + 1 And ToY% = SelY% - 1) Or (ToX% = SelX% - 1 And ToY% = SelY% + 1) Then Ok = 5
If (taking Or moving) And (ToX% = SelX% + 2 And ToY% = SelY% + 2) Or (ToX% = SelX% - 2 And ToY% = SelY% - 2) Or (ToX% = SelX% + 2 And ToY% = SelY% - 2) Or (ToX% = SelX% - 2 And ToY% = SelY% + 2) Then Ok = 5
If (taking Or moving) And (ToX% = SelX% + 3 And ToY% = SelY% + 3) Or (ToX% = SelX% - 3 And ToY% = SelY% - 3) Or (ToX% = SelX% + 3 And ToY% = SelY% - 3) Or (ToX% = SelX% - 3 And ToY% = SelY% + 3) Then Ok = 5
If (taking Or moving) And (ToX% = SelX% + 4 And ToY% = SelY% + 4) Or (ToX% = SelX% - 4 And ToY% = SelY% - 4) Or (ToX% = SelX% + 4 And ToY% = SelY% - 4) Or (ToX% = SelX% - 4 And ToY% = SelY% + 4) Then Ok = 5
If (taking Or moving) And (ToX% = SelX% + 5 And ToY% = SelY% + 5) Or (ToX% = SelX% - 5 And ToY% = SelY% - 5) Or (ToX% = SelX% + 5 And ToY% = SelY% - 5) Or (ToX% = SelX% - 5 And ToY% = SelY% + 5) Then Ok = 5
If (taking Or moving) And (ToX% = SelX% + 6 And ToY% = SelY% + 6) Or (ToX% = SelX% - 6 And ToY% = SelY% - 6) Or (ToX% = SelX% + 6 And ToY% = SelY% - 6) Or (ToX% = SelX% - 6 And ToY% = SelY% + 6) Then Ok = 5
If (taking Or moving) And (ToX% = SelX% + 7 And ToY% = SelY% + 7) Or (ToX% = SelX% - 7 And ToY% = SelY% - 7) Or (ToX% = SelX% + 7 And ToY% = SelY% - 7) Or (ToX% = SelX% - 7 And ToY% = SelY% + 7) Then Ok = 5
If (taking Or moving) And (ToX% = SelX% + 8 And ToY% = SelY% + 8) Or (ToX% = SelX% - 8 And ToY% = SelY% - 8) Or (ToX% = SelX% + 8 And ToY% = SelY% - 8) Or (ToX% = SelX% - 8 And ToY% = SelY% + 8) Then Ok = 5
If ToX% = SelX% Or ToY% = SelY% Then Ok = 0
If Ok = 0 Then Exit Sub
Case "Knight" 'Horsey
If (moving Or taking) And (ToX% = SelX% + 2 And ToY% = SelY% + 1) Then Ok = 6
If (moving Or taking) And (ToX% = SelX% + 1 And ToY% = SelY% + 2) Then Ok = 6
If (moving Or taking) And (ToX% = SelX% - 2 And ToY% = SelY% - 1) Then Ok = 6
If (moving Or taking) And (ToX% = SelX% - 1 And ToY% = SelY% - 2) Then Ok = 6
If (moving Or taking) And (ToX% = SelX% + 2 And ToY% = SelY% - 1) Then Ok = 6
If (moving Or taking) And (ToX% = SelX% + 1 And ToY% = SelY% - 2) Then Ok = 6
If (moving Or taking) And (ToX% = SelX% - 2 And ToY% = SelY% + 1) Then Ok = 6
If (moving Or taking) And (ToX% = SelX% - 1 And ToY% = SelY% + 2) Then Ok = 6
If Ok = 0 Then Exit Sub
Case "Castle" 'Castle
'Left
If (taking Or moving) And ToX% = SelX% - 1 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 2 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 3 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 4 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 5 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 6 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 7 And ToY% = SelY% Then Ok = 4
If (taking Or moving) And ToX% = SelX% - 8 And ToY% = SelY% Then Ok = 4
'Right
If (taking Or moving) And ToX% = SelX% + 1 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 2 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 3 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 4 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 5 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 6 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 7 And ToY% = SelY% Then Ok = 2
If (taking Or moving) And ToX% = SelX% + 8 And ToY% = SelY% Then Ok = 2
'Up
If (taking Or moving) And ToY% = SelY% - 1 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 2 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 3 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 4 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 5 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 6 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 7 And ToX% = SelX% Then Ok = 1
If (taking Or moving) And ToY% = SelY% - 8 And ToX% = SelX% Then Ok = 1
'Down
If (taking Or moving) And ToY% = SelY% + 1 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 2 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 3 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 4 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 5 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 6 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 7 And ToX% = SelX% Then Ok = 3
If (taking Or moving) And ToY% = SelY% + 8 And ToX% = SelX% Then Ok = 3
If Ok = 0 Then Exit Sub
End Select
8 'Check to see if we can move, without jumping a piece, based on OK values
Select Case Ok
Case 1 'Up
For y = SelY% - 1 To ToY% Step -1
If y < 0 Then Exit Sub
If moving And Board(SelX%, y) <> "" Then Exit Sub
If taking And (Board(SelX%, y) <> "" And Board(SelX%, y) <> Board(ToX%, ToY%)) Then Exit Sub
Next
Case 2 'Right
For x = SelX% + 1 To ToX% Step 1
If x < 0 Then Exit Sub
If moving And Board(x, SelY%) <> "" Then Exit Sub
If taking And (Board(x, SelY%) <> "" And Board(x, SelY%) <> Board(ToX%, ToY%)) Then Exit Sub
Next
Case 3 'Down
For y = SelY% + 1 To ToY% Step 1
If y < 0 Then Exit Sub
If moving And Board(SelX%, y) <> "" Then Exit Sub
If taking And (Board(SelX%, y) <> "" And Board(SelX%, y) <> Board(ToX%, ToY%)) Then Exit Sub
Next
Case 4 'Left
For x = SelX% - 1 To ToX% Step -1
If x < 0 Then Exit Sub
If moving And Board(x, SelY%) <> "" Then Exit Sub
If taking And (Board(x, SelY%) <> "" And Board(x, SelY%) <> Board(ToX%, ToY%)) Then Exit Sub
Next
Case 5 'Diagonal
If SelX% > ToX% And SelY% > ToY% Then 'Up Left
For x = SelX% - 1 To ToX% Step -1
For y = SelY% - 1 To ToY% Step -1
If SelX% - x = SelY% - y Then 'Only check if it's a diagonal
If moving And Board(x, y) <> "" Then Exit Sub
If taking And (Board(x, y) <> "" And Board(x, y) <> Board(ToX%, ToY%)) Then Exit Sub
End If
Next
Next
End If
If SelX% < ToX% And SelY% > ToY% Then 'Up Right
For x = SelX% + 1 To ToX%
For y = SelY% - 1 To ToY% Step -1
If SelX% - x = y - SelY% Then  'Only check if it's a diagonal
If moving And Board(x, y) <> "" Then Exit Sub
If taking And (Board(x, y) <> "" And Board(x, y) <> Board(ToX%, ToY%)) Then Exit Sub
End If
Next
Next
End If
If SelX% > ToX% And SelY% < ToY% Then 'Down Left
For x = SelX% - 1 To ToX% Step -1
For y = SelY% + 1 To ToY%
If x - SelX% = SelY% - y Then 'Only check if it's a diagonal
If moving And Board(x, y) <> "" Then Exit Sub
If taking And (Board(x, y) <> "" And Board(x, y) <> Board(ToX%, ToY%)) Then Exit Sub
End If
Next
Next
End If
If SelX% < ToX% And SelY% < ToY% Then 'Down Right
For x = SelX% + 1 To ToX%
For y = SelY% + 1 To ToY%
If SelX% - x = SelY% - y Then 'Only check if it's a diagonal
If moving And Board(x, y) <> "" Then Exit Sub
If taking And (Board(x, y) <> "" And Board(x, y) <> Board(ToX%, ToY%)) Then Exit Sub
End If
Next
Next
End If
Case 6 'It's a horsey & they're allowed to jump pieces!
End Select
'Successful Move!

'Check for CheckMate
If Board(ToX%, ToY%) <> "" Then
If Mode = "Online" And Turn = "W" And "King" = Left$(Board(ToX%, ToY%), Len(Board(ToX%, ToY%)) - 1) Then
MsgBox "You lose, due to the CHECK MATE rule", vbExclamation, "NSE ChessMASTER"
Ws.SendData "Vctry"
End If
If Mode = "Offline" And "King" = Left$(Board(ToX%, ToY%), Len(Board(ToX%, ToY%)) - 1) Then
If Turn = "B" Then i = MsgBox("You win. Play ChessMASTER again?", vbQuestion + vbYesNo, "NSE ChessMASTER")
If Turn = "W" Then i = MsgBox("ChessMASTER wins. Play ChessMASTER again?", vbQuestion + vbYesNo, "NSE ChessMASTER")
If i = vbYes Then
Mode = ""
For x = 0 To 8
For y = 0 To 8
Board(x, y) = ""
Next
Next
Command6_Click 'Click of the offline session button
Else
For x = 0 To 8
For y = 0 To 8
Board(x, y) = ""
Next
Next
Picture1.Cls
DrawBoard
Mode = ""
Turn = ""
Timer1.Enabled = False
End If
End If
End If

'Move
Board(ToX%, ToY%) = Board(SelX%, SelY%)
Board(SelX%, SelY%) = ""
Selected% = 0
DrawBoard

'Send Data to Opponent
If Mode = "Online" Then
Ws.SendData "move>" + Trim(Trim(Str(SelX%)) + Trim(Str(SelY%)) + Trim(Str(ToX%)) + Trim(Str(ToY%)))
End If

If Turn = "W" Then  'Change active player
Turn = "B"
Else
Turn = "W"
End If
End Sub

Function MoveAlright(sx, sy, tx, ty) As Boolean
'On Error Resume Next
If Board(sx, sy) = "" Then Exit Function
Colour$ = Right(Board(sx, sy), 1)
'See if were trying to take a piece
If Right(Board(sx, sy), 1) <> Right(Board(tx, ty), 1) And Right(Board(tx, ty), 1) <> "" Then taking = True
'See if were trying to take own piece!!
If Right(Board(sx, sy), 1) = Right(Board(tx, ty), 1) And Right(Board(tx, ty), 1) <> "" Then Exit Function
'See if moving (to blank space)
If Board(tx, ty) = "" Then moving = True

Select Case Mid(Board(sx, sy), 1, Len(Board(sx, sy)) - 1)
Case "Pawn"     'Prawn movement
If moving And (Board(tx, ty) = "" And ty = sy - 1 And Colour$ = "B" And sx = tx) Then Ok = 1    'White vertical
If moving And (Board(tx, ty) = "" And ty = sy + 1 And Colour$ = "W" And sx = tx) Then Ok = 3    'Black Vertical
If taking And ty = sy - 1 And Colour$ = "B" And (tx = sx - 1 Or tx = sx + 1) Then Ok = 5  'Black Up + Left/Right to White Taking a piece
If taking And ty = sy + 1 And Colour$ = "W" And (tx = sx - 1 Or tx = sx + 1) Then Ok = 5  'White Up + Left/Right to Black Taking a piece
If moving And (Colour$ = "B" And sy = 7 And ty = sy - 2) And tx = sx Then Ok = 1  'First move may be double BLACK
If moving And (Colour$ = "W" And sy = 2 And ty = sy + 2) And tx = sx Then Ok = 3  'First move may be double BLACK
If Ok = 0 Then Exit Function
Case "King"     'King movement
If (taking Or moving) And tx <= sx + 1 And tx >= sx - 1 And ty <= sy + 1 And ty >= sy - 1 Then Ok = 5 'Move in any direction by one
If Ok = 0 Then Exit Function
Case "Queen"
'Diagonal
If (taking Or moving) And sx - tx = sy - ty Or tx - tx = ty - sy Then Ok = 5
'Left
If (taking Or moving) And tx = sx - 1 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 2 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 3 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 4 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 5 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 6 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 7 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 8 And ty = sy Then Ok = 4
'Right
If (taking Or moving) And tx = sx + 1 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 2 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 3 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 4 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 5 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 6 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 7 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 8 And ty = sy Then Ok = 2
'Up
If (taking Or moving) And ty = sy - 1 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 2 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 3 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 4 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 5 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 6 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 7 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 8 And tx = sx Then Ok = 1
'Down
If (taking Or moving) And ty = sy + 1 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 2 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 3 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 4 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 5 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 6 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 7 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 8 And tx = sx Then Ok = 3
If Ok = 0 Then Exit Function
Case "Bishop"
If (taking Or moving) And (tx = sx + 1 And ty = sy + 1) Or (tx = sx - 1 And ty = sy - 1) Or (tx = sx + 1 And ty = sy - 1) Or (tx = sx - 1 And ty = sy + 1) Then Ok = 5
If (taking Or moving) And (tx = sx + 2 And ty = sy + 2) Or (tx = sx - 2 And ty = sy - 2) Or (tx = sx + 2 And ty = sy - 2) Or (tx = sx - 2 And ty = sy + 2) Then Ok = 5
If (taking Or moving) And (tx = sx + 3 And ty = sy + 3) Or (tx = sx - 3 And ty = sy - 3) Or (tx = sx + 3 And ty = sy - 3) Or (tx = sx - 3 And ty = sy + 3) Then Ok = 5
If (taking Or moving) And (tx = sx + 4 And ty = sy + 4) Or (tx = sx - 4 And ty = sy - 4) Or (tx = sx + 4 And ty = sy - 4) Or (tx = sx - 4 And ty = sy + 4) Then Ok = 5
If (taking Or moving) And (tx = sx + 5 And ty = sy + 5) Or (tx = sx - 5 And ty = sy - 5) Or (tx = sx + 5 And ty = sy - 5) Or (tx = sx - 5 And ty = sy + 5) Then Ok = 5
If (taking Or moving) And (tx = sx + 6 And ty = sy + 6) Or (tx = sx - 6 And ty = sy - 6) Or (tx = sx + 6 And ty = sy - 6) Or (tx = sx - 6 And ty = sy + 6) Then Ok = 5
If (taking Or moving) And (tx = sx + 7 And ty = sy + 7) Or (tx = sx - 7 And ty = sy - 7) Or (tx = sx + 7 And ty = sy - 7) Or (tx = sx - 7 And ty = sy + 7) Then Ok = 5
If (taking Or moving) And (tx = sx + 8 And ty = sy + 8) Or (tx = sx - 8 And ty = sy - 8) Or (tx = sx + 8 And ty = sy - 8) Or (tx = sx - 8 And ty = sy + 8) Then Ok = 5
If tx = sx Or ty = sy Then Ok = 0
If Ok = 0 Then Exit Function
Case "Knight" 'Horsey
If (moving Or taking) And (tx = sx + 2 And ty = sy + 1) Then Ok = 6
If (moving Or taking) And (tx = sx + 1 And ty = sy + 2) Then Ok = 6
If (moving Or taking) And (tx = sx - 2 And ty = sy - 1) Then Ok = 6
If (moving Or taking) And (tx = sx - 1 And ty = sy - 2) Then Ok = 6
If (moving Or taking) And (tx = sx + 2 And ty = sy - 1) Then Ok = 6
If (moving Or taking) And (tx = sx + 1 And ty = sy - 2) Then Ok = 6
If (moving Or taking) And (tx = sx - 2 And ty = sy + 1) Then Ok = 6
If (moving Or taking) And (tx = sx - 1 And ty = sy + 2) Then Ok = 6
If Ok = 0 Then Exit Function
Case "Castle" 'Castle
'Left
If (taking Or moving) And tx = sx - 1 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 2 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 3 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 4 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 5 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 6 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 7 And ty = sy Then Ok = 4
If (taking Or moving) And tx = sx - 8 And ty = sy Then Ok = 4
'Right
If (taking Or moving) And tx = sx + 1 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 2 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 3 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 4 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 5 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 6 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 7 And ty = sy Then Ok = 2
If (taking Or moving) And tx = sx + 8 And ty = sy Then Ok = 2
'Up
If (taking Or moving) And ty = sy - 1 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 2 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 3 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 4 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 5 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 6 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 7 And tx = sx Then Ok = 1
If (taking Or moving) And ty = sy - 8 And tx = sx Then Ok = 1
'Down
If (taking Or moving) And ty = sy + 1 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 2 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 3 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 4 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 5 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 6 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 7 And tx = sx Then Ok = 3
If (taking Or moving) And ty = sy + 8 And tx = sx Then Ok = 3
If Ok = 0 Then Exit Function
End Select
8 'Check to see if we can move, without jumping a piece, based on OK values
Select Case Ok
Case 1 'Up
For y = sy - 1 To ty Step -1
If y < 0 Then Exit Function
If moving And Board(sx, y) <> "" Then Exit Function
If taking And (Board(sx, y) <> "" And Board(sx, y) <> Board(tx, ty)) Then Exit Function
Next
Case 2 'Right
For x = sx + 1 To tx Step 1
If x < 0 Then Exit Function
If moving And Board(x, sy) <> "" Then Exit Function
If taking And (Board(x, sy) <> "" And Board(x, sy) <> Board(tx, ty)) Then Exit Function
Next
Case 3 'Down
For y = sy + 1 To ty Step 1
If y < 0 Then Exit Function
If moving And Board(sx, y) <> "" Then Exit Function
If taking And (Board(sx, y) <> "" And Board(sx, y) <> Board(tx, ty)) Then Exit Function
Next
Case 4 'Left
For x = sx - 1 To tx Step -1
If x < 0 Then Exit Function
If moving And Board(x, sy) <> "" Then Exit Function
If taking And (Board(x, sy) <> "" And Board(x, sy) <> Board(tx, ty)) Then Exit Function
Next
Case 5 'Diagonal
If sx > tx And sy > ty Then 'Up Left
For x = sx - 1 To tx Step -1
For y = sy - 1 To ty Step -1
If sx - x = sy - y Then 'Only check if it's a diagonal
If moving And Board(x, y) <> "" Then Exit Function
If taking And (Board(x, y) <> "" And Board(x, y) <> Board(tx, ty)) Then Exit Function
End If
Next
Next
End If
If sx < tx And sy > ty Then 'Up Right
For x = sx + 1 To tx
For y = sy - 1 To ty Step -1
If sx - x = y - sy Then  'Only check if it's a diagonal
If moving And Board(x, y) <> "" Then Exit Function
If taking And (Board(x, y) <> "" And Board(x, y) <> Board(tx, ty)) Then Exit Function
End If
Next
Next
End If
If sx > tx And sy < ty Then 'Down Left
For x = sx - 1 To tx Step -1
For y = sy + 1 To ty
If x - sx = sy - y Then 'Only check if it's a diagonal
If moving And Board(x, y) <> "" Then Exit Function
If taking And (Board(x, y) <> "" And Board(x, y) <> Board(tx, ty)) Then Exit Function
End If
Next
Next
End If
If sx < tx And sy < ty Then 'Down Right
For x = sx + 1 To tx
For y = sy + 1 To ty
If sx - x = sy - y Then 'Only check if it's a diagonal
If moving And Board(x, y) <> "" Then Exit Function
If taking And (Board(x, y) <> "" And Board(x, y) <> Board(tx, ty)) Then Exit Function
End If
Next
Next
End If
Case 6 'It's a horsey & they're allowed to jump pieces!
End Select

'Successful Move!
MoveAlright = True
End Function

Function QualityOfTarget(x, y)
If Board(x, y) = "" Then
QualityOfTarget = 0
Exit Function
End If

Select Case Mid(Board(x, y), 1, Len(Board(x, y)) - 1)
Case "Pawn"
QualityOfTarget = 4
Case "Castle"
QualityOfTarget = 9
Case "Knight"
QualityOfTarget = 9
Case "Bishop"
QualityOfTarget = 9
Case "Queen"
QualityOfTarget = 14
Case "King"
QualityOfTarget = 100
End Select
End Function

Function CostOfLoss(x, y)
If Board(x, y) = "" Then Exit Function
Select Case Mid(Board(x, y), 1, Len(Board(x, y)) - 1)
Case "Pawn"
CostOfLoss = 5
Case "Castle"
CostOfLoss = 10
Case "Knight"
CostOfLoss = 10
Case "Bishop"
CostOfLoss = 10
Case "Queen"
CostOfLoss = 15
Case "King"
CostOfLoss = 100
End Select
End Function

