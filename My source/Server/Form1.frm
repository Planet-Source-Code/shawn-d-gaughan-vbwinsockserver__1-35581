VERSION 5.00
Begin VB.Form display 
   Caption         =   "Server Display"
   ClientHeight    =   6750
   ClientLeft      =   2880
   ClientTop       =   3555
   ClientWidth     =   13470
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   13470
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9840
      Top             =   1200
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   13470
      TabIndex        =   1
      Top             =   0
      Width           =   13470
      Begin VB.PictureBox table 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -120
         ScaleHeight     =   615
         ScaleWidth      =   13395
         TabIndex        =   2
         Top             =   0
         Width           =   13395
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
            Left            =   120
            Picture         =   "Form1.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   615
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
            Left            =   720
            Picture         =   "Form1.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   0
            Width           =   615
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
            Left            =   1320
            Picture         =   "Form1.frx":0CC6
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   0
            Width           =   735
         End
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
            Left            =   2040
            Picture         =   "Form1.frx":1108
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   0
            Width           =   615
         End
         Begin VB.CommandButton Command5 
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
            Left            =   2640
            Picture         =   "Form1.frx":154A
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton Command6 
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
            Left            =   4080
            Picture         =   "Form1.frx":198C
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton Command7 
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
            Left            =   3360
            Picture         =   "Form1.frx":1DCE
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   0
            Width           =   735
         End
      End
   End
   Begin VB.TextBox stat 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   6615
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   13455
   End
   Begin VB.Line Line2 
      BorderWidth     =   10
      X1              =   0
      X2              =   4920
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

stat = ""
stat = "    ---> Max Clients:  " & maxcl & Chr(13) & Chr(10)

For i = 0 To (maxcl - 1)
stat = stat & "  --> Socket:  " & i & "  Using port:  " & clt(i).port & "  Connected:  " & clt(i).state & Chr(13) & Chr(10)
Next i


End Sub

Private Sub Command2_Click()

Dim ccount As Integer
ccount = 0
stat = ""

For i = 0 To (maxcl - 1)

If clt(i).state = True Then
ccount = ccount + 1
stat = stat & "**********************************************" & Chr(13) & Chr(10)
stat = stat & "Socket:  " & i & "  Port: " & clt(i).port & Chr(13) & Chr(10)
stat = stat & "Username:  " & clt(i).username & "  ip: " & clt(i).ip & Chr(13) & Chr(10)
End If

Next i

stat = Chr(13) & Chr(10) & "  <--- Current User Online: " & ccount & Chr(13) & Chr(10) & Chr(13) & Chr(10) & stat



End Sub

Private Sub Command3_Click()


Dim ccount As Integer
Dim blen As Integer
ccount = 0
stat = ""

For i = 0 To (maxcl - 1)
If clt(i).state = True And clt(i).buffer <> "" Then
ccount = ccount + 1
stat = stat & "**********************************************" & Chr(13) & Chr(10)
stat = stat & "Socket: " & i & " Port:" & clt(i).port & Chr(13) & Chr(10)
stat = stat & "Username: " & clt(i).username & " ip: " & clt(i).ip & Chr(13) & Chr(10)
stat = stat & "Buffer: " & clt(i).buffer & Chr(13) & Chr(10)
blen = blen + Len(clt(i).buffer)
End If

Next i

stat = Chr(13) & Chr(10) & "  <--- Bufffers Bites: " & blen & " Buffers In Use: " & ccount & Chr(13) & Chr(10) & Chr(13) & Chr(10) & stat



End Sub

Private Sub Command4_Click()

stat = slog

End Sub

Private Sub Command5_Click()
stat = ""


End Sub

Private Sub Command6_Click()
display.Hide
End Sub

Private Sub Command7_Click()

stat = ""

End Sub



Private Sub Form_Resize()
On Error Resume Next
stat.Width = Me.Width - 100
stat.Height = Me.Height - 1100



End Sub


