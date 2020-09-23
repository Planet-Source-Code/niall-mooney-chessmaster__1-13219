VERSION 5.00
Begin VB.Form frmEnter 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ni-Star Enterprises - ChessMASTER"
   ClientHeight    =   1800
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4020
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Enter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1063.499
   ScaleMode       =   0  'User
   ScaleWidth      =   3774.562
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Caption         =   "Logon to ChessMASTER as a client to connect to a server?"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Caption         =   "Logon to ChessMASTER as a server or for Offline play?"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Value           =   -1  'True
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Do you wish to:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.ForeColor = &HFF8080
End Sub

Private Sub Label2_Click()
If Option1(0).Value = True Then
Unload Me
frmServer.Show
Else
Unload Me
frmClient.Show
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.ForeColor = &HFFFF&
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Label2_Click
End Sub

Private Sub Option1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.ForeColor = &HFF8080
End Sub
