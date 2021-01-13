Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Dim s$

Sub Main()
On Error Resume Next
    Do
        Delay 300
        s = Clipboard.GetText
        If s <> "" Then
            If InStr(s, vbCrLf) = 0 And InStr(s, vbLf) > 0 Then
                s = Replace(s, vbLf, vbCrLf)
                Clipboard.Clear
                Clipboard.SetText s
            End If
        End If
    Loop
End Sub

'µ¥Î»ÎªºÁÃë
Public Sub Delay(ByVal MS As Long)
    Dim T As Long
    T = timeGetTime
    While timeGetTime - T < MS
        DoEvents
    Wend
End Sub
