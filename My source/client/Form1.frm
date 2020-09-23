VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   3615
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1920
      Width           =   8655
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1680
      MaxLength       =   500
      TabIndex        =   4
      Top             =   1200
      Width           =   6615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "send"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Text            =   "66.32.234.141"
      Top             =   360
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2280
      TabIndex        =   1
      Text            =   "2600"
      Top             =   360
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8400
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "connect"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Winsock1.State = sckConnected Then
Do
Winsock1.Close
Winsock1.Close
Winsock1.Close
Loop Until Winsock1.State = sckClosed
End If


Winsock1.RemoteHost = Text2
Winsock1.RemotePort = Text1

Winsock1.Connect

Do Until Winsock1.State = sckConnected
DoEvents: DoEvents: DoEvents: DoEvents
Loop

MsgBox "Connected"

End Sub

Private Sub Command2_Click()
On Error Resume Next
'Convert simple

Dim data As String
cc = Len(Text3)

'check
If Len(cc) = 1 Then data = "00" & Len(Text3)
If Len(cc) = 2 Then data = "0" & Len(Text3)
bla = data & "TK" & Text3

Winsock1.SendData bla
Text3 = ""

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim bla As String
Winsock1.GetData bla
Text3.SetFocus
Text4 = Text4 & bla & Chr(13) & Chr(10)
Text4.SelStart = Len(Text4)
End Sub

