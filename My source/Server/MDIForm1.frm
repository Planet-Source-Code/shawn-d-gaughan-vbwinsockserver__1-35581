VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm main 
   BackColor       =   &H8000000C&
   Caption         =   "Server "
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock server 
      Index           =   9999
      Left            =   10080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9120
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1650
      Left            =   0
      ScaleHeight     =   1650
      ScaleWidth      =   12585
      TabIndex        =   0
      Top             =   0
      Width           =   12585
      Begin VB.Frame ics 
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   1815
         Left            =   10800
         TabIndex        =   10
         Top             =   120
         Width           =   7935
         Begin MSComctlLib.ProgressBar uprog 
            DragMode        =   1  'Automatic
            Height          =   375
            Left            =   360
            TabIndex        =   11
            Top             =   480
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Max             =   500
         End
         Begin VB.Label uos 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0 / 500 Users online."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   480
            TabIndex        =   14
            Top             =   120
            Width           =   6855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Banned Ip Addresses: 0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   360
            TabIndex        =   13
            Top             =   960
            Width           =   7455
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "User Names Database: 0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   585
            Left            =   360
            TabIndex        =   12
            Top             =   1200
            Width           =   7290
         End
      End
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   915
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   3075
         TabIndex        =   5
         Top             =   960
         Width           =   3075
         Begin VB.CommandButton Command4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   720
            Picture         =   "MDIForm1.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            Picture         =   "MDIForm1.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton Command2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1440
            Picture         =   "MDIForm1.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2160
            Picture         =   "MDIForm1.frx":0CC6
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   0
            Width           =   735
         End
      End
      Begin MSWinsockLib.Winsock wintest 
         Left            =   9480
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   840
         Width           =   5535
      End
      Begin VB.Label ls 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   360
         Width           =   5895
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Line Line3 
         X1              =   4080
         X2              =   4080
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   4080
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4080
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label sd 
         Caption         =   "Server Date: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label st 
         Caption         =   "Server Time: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   120
         Width           =   2895
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
display.Show

End Sub







Private Sub MDIForm_Load()

ls = "Local Server Address: " & wintest.LocalIP

'Server Writting by shawn/Forbidden

'Veriables configure
sport = 1 'starting port
maxcl = 500 'Maximum clients
plis = 2600 ' Server listing port

X = sport - 1
b = -1
ucount = 0

'Check port range
For i = sport To (sport + maxcl)

'pointer
b = b + 1

'Close port
Do
wintest.Close
Loop Until wintest.state = sckClosed
GoTo ski:
k:
Do
wintest.Close
Loop Until wintest.state = sckClosed
'Error Log error
elog = elog & "ERROR: " & Err.Description & " " & wintest.LocalPort & "- Port Changed -  " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
Err.Clear
Resume Next
ski:
On Error GoTo k:
X = X + 1
wintest.LocalPort = X
wintest.Listen

'Record info to memory
clt(b).state = False
clt(b).port = X
Next i

'Close port
Do
wintest.Close
Loop Until wintest.state = sckClosed

On Error GoTo ehan:
'Open listing port
wintest.LocalPort = plis
wintest.Listen

If elog <> "" Then
MsgBox "The Server was started." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Server Recovered from the following conflicts" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & elog, vbInformation, "Server Conflicts"
End If

lp = "Server Listing On Port: " & plis

slog = " >> Server started on port: " & plis & " Time: " & Time & " Date: " & Date & Chr(13) & Chr(10)

Exit Sub


'Error handeling
ehan:
MsgBox "Listing Port In USE Please free up the listing port: " & plis & " or reconfigure the server, Server may not be started", vbCritical
End
Exit Sub

End Sub






Private Sub server_Close(Index As Integer)

slog = slog & Time & ": >> Client DisConnected Normal: " & clt(Index).ip & Chr(13) & Chr(10)

'clear info unload buffer
clt(Index).username = ""
'clt(Index).room() = Null
clt(Index).password = ""
clt(Index).ip = ""
clt(Index).buffer = ""
clt(Index).state = False
Unload server(Index)
ucount = ucount - 1
End Sub



Private Sub server_DataArrival(Index As Integer, ByVal bytesTotal As Long)

'prevents exploiting
If Index = 9999 Then Exit Sub

Dim str As String
server(Index).GetData str

clt(Index).buffer = clt(Index).buffer & str

'dims
Dim bites As Integer
Dim cmb As Integer
Dim data As String
Dim X As Integer
Dim ck
Dim lbuffer As String

lbuffer = clt(Index).buffer

'Quick Check
If Len(lbuffer) <= 5 Then Exit Sub
'Check first 3 digits
ck = Mid(lbuffer, 1, 3)
cdd1 = Mid(lbuffer, 4, 1)
cdd2 = Mid(lbuffer, 5, 1)

If Not IsNumeric(ck) Or (IsNumeric(cdd1) Or IsNumeric(cdd2)) Then
screwthis:
'buffer errored spected user violation
Timeout 0.1
If clt(Index).state = True Then
slog = slog & Time & ": " & " >> Client DisConnected Fishy: " & clt(Index).ip & Chr(13) & Chr(10)
ucount = ucount - 1
End If

'force boot
'clear info unload buffer
clt(Index).username = ""
'clt(Index).room() = Null
clt(Index).password = ""
clt(Index).ip = ""
clt(Index).buffer = ""
clt(Index).state = False
On Error Resume Next
Unload server(Index)
Timeout 0.1
Exit Sub
End If

'convert
bites = Mid(lbuffer, 1, 3)
cmd = Mid(lbuffer, 4, 2)
'Define data
data = Mid(lbuffer, 6, bites)
'Valadate data

If Len(data) = 0 Or Len(data) > 999 Or Len(lbuffer) >= 999 Then GoTo screwthis:

If bites <= Len(data) Then


'ADD Data to list
Dim text2 As String
text2 = Chr(13) & Chr(10)
text2 = text2 & clt(Index).ip & " says " & Chr(13) & Chr(10)
text2 = text2 & "Bites: " & bites & Chr(13) & Chr(10)
text2 = text2 & "CMD: " & cmd & Chr(13) & Chr(10)
text2 = text2 & "Data: " & data & Chr(13) & Chr(10)

'send to all
On Error Resume Next
For n = 0 To 500
If clt(n).state = True Then
Timeout 0.1
server(n).SendData code(text2, "SB")
End If
Next n


'MsgBox text2

'Remove from buffer
X = 6 + Len(data)
lbuffer = Mid(lbuffer, X)

clt(Index).buffer = lbuffer

End If



End Sub




Private Sub server_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

'clear info unload buffer
clt(Index).username = ""
'clt(Index).room() = Null
clt(Index).password = ""
clt(Index).ip = ""
clt(Index).buffer = ""
clt(Index).state = False
Unload server(Index)

End Sub

Private Sub Timer1_Timer()
st = "Server Time: " & Time
sd = "Server Date: " & Date

uos = ucount & " / " & maxcl & " Users online."
uprog.Value = ucount
uprog.Max = maxcl

End Sub

Private Sub wintest_ConnectionRequest(ByVal requestID As Long)


'IP Denied ?
'rip = "127.0.0.1"
If wintest.RemoteHostIP = rip Then Exit Sub



'Load new winsock
For i = 0 To (maxcl - 1)
On Error Resume Next
'check socket if ready then load up
If clt(i).state = False Then
clt(i).state = True
clt(i).ip = wintest.RemoteHostIP
slog = slog & Time & ": " & "^Connect ReQuest: " & wintest.RemoteHostIP & Chr(13) & Chr(10)
Load server(i)
server(i).Accept requestID
slog = slog & Time & ": " & "^ReQuest Connected: " & server(i).RemoteHostIP & " " & server(i).LocalPort & Chr(13) & Chr(10)
Timeout 1

'welcome message
ucount = ucount + 1
server(i).SendData code("********* Welcome USers ***********", "WK")

slog = slog & Time & ": " & "*Welcome Message: " & server(i).RemoteHostIP & Chr(13) & Chr(10)


Exit Sub
End If

Next i

End Sub



