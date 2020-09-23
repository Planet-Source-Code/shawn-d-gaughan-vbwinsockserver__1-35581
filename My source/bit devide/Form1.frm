VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Streaming Buffer"
   ClientHeight    =   3690
   ClientLeft      =   5310
   ClientTop       =   2925
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   7635
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7080
      Top             =   600
   End
   Begin VB.TextBox Text2 
      Height          =   2295
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Text            =   "00201HI"
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Public Function changeb(button As CommandButton, cap As String)

button.Caption = cap

End Function






Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()


'dims
Dim bites As Integer
Dim cmb As Integer
Dim data As String
Dim x As Integer
Dim ck
'Quick Check
If Len(Text1) <= 5 Then Exit Sub
'Check first 3 digits
ck = Mid(Text1, 1, 5)
If Not IsNumeric(ck) Then
MsgBox "Fetal ERROR In Buffer, Client can not continue", vbCritical
End
Exit Sub
End If
'convert
bites = Mid(Text1, 1, 3)
cmd = Mid(Text1, 4, 2)
'Define data
data = Mid(Text1, 6, bites)
'Valadate data
If bites <= Len(data) Then
'ADD Data to list
Text2 = Text2 & "Bites: " & bites & Chr(13) & Chr(10)
Text2 = Text2 & "CMD: " & cmd & Chr(13) & Chr(10)
Text2 = Text2 & "Data: " & data & Chr(13) & Chr(10)
'Scrool
Text2.SelStart = Len(Text2)
'Remove from buffer
x = 6 + Len(data)
Text1 = Mid(Text1, x)
End If


End Sub
