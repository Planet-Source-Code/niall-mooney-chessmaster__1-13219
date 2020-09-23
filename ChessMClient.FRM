VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmClient 
   Caption         =   "Ni-Star Enterprises - ChessMASTER Client Mode"
   ClientHeight    =   7620
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ChessMClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   553
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   12
      Left            =   2100
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
      TabIndex        =   35
      Top             =   6120
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send Popup"
      Height          =   255
      Left            =   6960
      TabIndex        =   34
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Board"
      Height          =   1215
      Left            =   5880
      TabIndex        =   30
      Top             =   6360
      Width           =   2415
      Begin VB.OptionButton Option1 
         Caption         =   "Alternative Board"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Original but better"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Original and best"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.TextBox Terminal 
      Height          =   2595
      Left            =   5835
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   2505
      Width           =   2430
   End
   Begin VB.TextBox Handle 
      Height          =   285
      Left            =   5820
      TabIndex        =   24
      Top             =   2145
      Width           =   2445
   End
   Begin VB.PictureBox Piece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   11
      Left            =   2100
      Picture         =   "ChessMClient.frx":030A
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
      Picture         =   "ChessMClient.frx":0798
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
      Picture         =   "ChessMClient.frx":0C59
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
      Picture         =   "ChessMClient.frx":11DE
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
      Picture         =   "ChessMClient.frx":16A4
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
      Picture         =   "ChessMClient.frx":1B82
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
      Picture         =   "ChessMClient.frx":2048
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
      Picture         =   "ChessMClient.frx":24D6
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
      Picture         =   "ChessMClient.frx":2997
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
      Picture         =   "ChessMClient.frx":2F1C
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
      Picture         =   "ChessMClient.frx":33E2
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
      Picture         =   "ChessMClient.frx":38C0
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
      Top             =   5775
      Width           =   1035
   End
   Begin VB.TextBox Chattxt 
      Height          =   585
      Left            =   5835
      TabIndex        =   8
      Top             =   5145
      Width           =   2430
   End
   Begin VB.Frame Frame1 
      Caption         =   "WinSocket Control"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton Command5 
         Caption         =   "Save As..."
         Height          =   255
         Left            =   3120
         TabIndex        =   37
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Resign/Concede"
         Height          =   255
         Left            =   4320
         TabIndex        =   36
         Top             =   1440
         Width           =   1815
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   1665
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "NSE ChessMASTER"
      End
      Begin MSWinsockLib.Winsock Ws 
         Left            =   1560
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   499
         LocalPort       =   500
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Refresh Board"
         Height          =   255
         Left            =   4320
         TabIndex        =   28
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Send..."
         Height          =   255
         Left            =   3120
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Chat..."
         Height          =   255
         Left            =   4320
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Disconnect"
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Connect"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   5280
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox IpConnecttxt 
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   3015
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
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Your Handle is:"
      Height          =   255
      Left            =   5820
      TabIndex        =   25
      Top             =   1905
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
      Height          =   1710
      Left            =   6450
      TabIndex        =   11
      Top             =   105
      Width           =   1800
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu MnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSave 
         Caption         =   "&Save As"
      End
      Begin VB.Menu MnuRetire 
         Caption         =   "&Resign"
      End
      Begin VB.Menu MnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuFeature 
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
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Board(8, 8) As String
Dim Mode As String, Selected%, SelX%, SelY%, ToX%, ToY%
Dim Turn As String, TransColour As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Sub UpdateConnectionStats()
IPlbl.Caption = "Your IP: " & Ws.LocalIP
Nmlbl.Caption = "Your Name: " & Ws.LocalHostName
Oiplbl.Caption = "Opponent IP: " & Ws.RemoteHostIP
oNmlbl.Caption = "Opponent Name: " & Ws.RemoteHost
Handle.Text = Ws.LocalHostName
End Sub

Private Sub Chattxt_Change()
Command7.Default = True
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Ws.State <> sckClosed Then
MsgBox "Winsock is repoting sckOpen status. Please disconnect!", vbCritical, "NSE ChessMASTER Online"
Exit Sub
End If
Ws.Connect IpConnecttxt.Text, 0
UpdateConnectionStats

If Err.Description <> "" Then MsgBox Err.Description, vbCritical, "NSE ChessMaster"

Mode = "Online"
Turn = "W"
InitiateOnlineMode
End Sub

Private Sub Command10_Click()
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

Private Sub Command3_Click()
Ws.Close
End Sub

Private Sub Command4_Click()
If Mode = "" Then Exit Sub

i = MsgBox("Are you sure you want to resign", vbQuestion + vbYesNo, "NSE ChessMASTER")
If i = vbYes Then
Mode = ""
For x = 0 To 8
For y = 0 To 8
Board(x, y) = ""
Next
Next
Picture1.Cls
MsgBox "You lose", vbExclamation, "NSE ChessMASTER"
Ws.SendData "Vctry"
End If
End Sub

Private Sub Command5_Click()
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

Private Sub Command6_Click()
On Error Resume Next
If MsgBox("Start a new Offline Mode session", vbMsgBoxSetForeground + vbQuestion + vbYesNo, "NSE ChessMASTER") = vbNo Then Exit Sub
initBoard
DrawBoard
Ws.RemoteHost = "N.I.A.L.L."
Mode = "Offline"
Turn = "W"
Timer1.Enabled = True
End Sub

Private Sub Command7_Click()
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
CD.InitDir = App.Path
Me.Show
TransColour = vbWhite
Option1_Click 1
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
UpdateConnectionStats
End Sub

Private Sub MnuAbout_Click()
frmAbout.Show
End Sub

Private Sub MnuBlue_Click()
TransColour = vbBlue
DrawBoard
End Sub

Private Sub MnuChat_Click()
Command9_Click
End Sub

Private Sub MnuConnect_Click()
Command1_Click
End Sub

Private Sub MnuCyan_Click()
TransColour = vbCyan
DrawBoard
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

Private Sub MnuRed_Click()
TransColour = vbRed
DrawBoard
End Sub

Private Sub MnuRetire_Click()
Command4_Click
End Sub

Private Sub MnuSave_Click()
Command5_Click
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

If Selected% = 1 Then Coords.Caption = Coords.Caption & vbNewLine & Mid(Board(SelX%, SelY%), 1, Len(Board(SelX%, SelY%)) - 1) & " @ " & SelX% & ", " & SelY% & " to " & ToX% & ", " & ToY%
If Selected% = 1 And MoveAlright(SelX%, SelY%, ToX%, ToY%) Then Coords.Caption = Coords.Caption & vbNewLine & "-LEGAL MOVE-"

If Turn = "W" Then Coords.Caption = Coords.Caption & vbNewLine & "Waiting For " & Ws.RemoteHost
If Turn = "B" Then Coords.Caption = Coords.Caption & vbNewLine & "Your Turn, " & Ws.LocalHostName
End Sub

Private Sub Timer1_Timer()
Dim Goodness
Dim Xes
Dim Yes
If Turn = "W" Then
For x = 1 To 8
For y = 1 To 8
If Right(Board(x, y), 1) = "W" Then 'It's a moveable peice
For X2 = 0 To 8 'get any good enemies with it?
For Y2 = 1 To 8
If Right(Board(X2, Y2), 1) = "B" Then
If MoveAlright(x, y, X2, Y2) Then
End If
End If
Next
Next
End If
Next
Next
End If
End Sub

Private Sub Ws_Connect()
Me.Caption = "Ni-Star Enterprises - ChessMASTER Client Mode CONNECTED"
UpdateConnectionStats
End Sub

Private Sub Ws_DataArrival(ByVal bytesTotal As Long)
Ws.GetData DataReceived, vbString, 90 '86 bytes for the board

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

If Left$(DataReceived, 5) = "load>" Then 'It's a new board
i = 6
For x = 0 To 8
For y = 0 To 8
If Mid(DataReceived, 5, 1) = "W" Then 'Assign Turn
Turn = "B"
Else
Turn = "W"
End If
Board(x, y) = Val(Mid(DataReceived, i, 1))
i = i + 8 'Move on 8 Bytes in the board string to load
Next
Next
MsgBox "New board has been recieved from " & Ws.RemoteHost & " & sucessfully loaded.", vbInformation, "NSE ChessMASTER"
DrawBoard
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
If (taking Or moving) And SelX% - ToX% = SelY% - ToY% Then Ok = 5
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
If Turn = "W" And "King" = Left$(Board(ToX%, ToY%), Len(Board(ToX%, ToY%)) - 1) Then
MsgBox "You lose, due to the CHECK MATE rule", vbExclamation, "NSE ChessMASTER"
Ws.SendData "Vctry"
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
If (taking Or moving) And sx - tx = sy - ty Then Ok = 5
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

