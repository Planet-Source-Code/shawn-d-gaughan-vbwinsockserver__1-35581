Attribute VB_Name = "server"
'Define types

Public Type bla
buffer As String
port As Integer
state As Boolean
password As String
username As String
ip As String
room() As String
End Type


'Log data
Public slog As String


'defined ports
Public ucount As Integer
Public clt(500) As bla
Public elog As String
Public maxcl As Integer


Function code(dat As String, cmd As String)

'Convert
If Len(dat) > 999 Or Len(cmd) <> 2 Then
MsgBox "function can't convert"
Exit Function
End If


Dim data As String
cc = Len(dat)

'check
If Len(cc) = 1 Then data = "00" & Len(dat)
If Len(cc) = 2 Then data = "0" & Len(dat)

bla = data & cmd & dat

code = bla

End Function


Sub Timeout(duration)
Dim starttime, X
startime = Timer
Do While Timer - startime < duration
X = DoEvents()
Loop
End Sub



