Attribute VB_Name = "Module4"
Option Explicit

Public Sub SET_COLORTEMP(colorT As Long)
Dim SendDataBuf(0 To 11) As Byte
Dim value As Byte
'55  02  01  XX  00  00  00  00  00  00      FE
Select Case colorT
  Case 13000
     value = CByte(IsCool_1ModeIndex)
  Case 11000
     value = CByte(IsNormalModeIndex)
  Case 9300
     value = CByte(IsWarm_1ModeIndex)
End Select
SendDataBuf(0) = &H55
SendDataBuf(1) = &H2
SendDataBuf(2) = &H1
SendDataBuf(3) = value
SendDataBuf(4) = &H0
SendDataBuf(5) = &H0
SendDataBuf(6) = &H0
SendDataBuf(7) = &H0
SendDataBuf(8) = &H0
SendDataBuf(9) = &H0
SendDataBuf(10) = chksumSend(SendDataBuf)
SendDataBuf(11) = &HFE

Form1.MSComm1.Output = SendDataBuf
DelayMS 500
End Sub

Public Sub Save_Cool1()
Dim SendDataBuf(0 To 11) As Byte

SendDataBuf(0) = &H55
SendDataBuf(1) = &H16
SendDataBuf(2) = &H1
SendDataBuf(3) = &H3
SendDataBuf(4) = &H0
SendDataBuf(5) = &H0
SendDataBuf(6) = &H0
SendDataBuf(7) = &H0
SendDataBuf(8) = &H0
SendDataBuf(9) = &H0
SendDataBuf(10) = &HE6
SendDataBuf(11) = &HFE

Form1.MSComm1.Output = SendDataBuf
DelayMS 800
End Sub
Public Sub Save_Normal()
Dim SendDataBuf(0 To 11) As Byte
SendDataBuf(0) = &H55
SendDataBuf(1) = &H16
SendDataBuf(2) = &H1
SendDataBuf(3) = &H3
SendDataBuf(4) = &H0
SendDataBuf(5) = &H0
SendDataBuf(6) = &H0
SendDataBuf(7) = &H0
SendDataBuf(8) = &H0
SendDataBuf(9) = &H0
SendDataBuf(10) = &HE6
SendDataBuf(11) = &HFE

Form1.MSComm1.Output = SendDataBuf
DelayMS 800
End Sub
Public Sub Save_Warm1()
Dim SendDataBuf(0 To 11) As Byte
SendDataBuf(0) = &H55
SendDataBuf(1) = &H16
SendDataBuf(2) = &H1
SendDataBuf(3) = &H3
SendDataBuf(4) = &H0
SendDataBuf(5) = &H0
SendDataBuf(6) = &H0
SendDataBuf(7) = &H0
SendDataBuf(8) = &H0
SendDataBuf(9) = &H0
SendDataBuf(10) = &HE6
SendDataBuf(11) = &HFE

Form1.MSComm1.Output = SendDataBuf
DelayMS 800
End Sub

Public Sub Save_WhiteBlance(colorT As Long)
Dim SendDataBuf(0 To 11) As Byte
Dim value As Byte
'55  16  01  XX  00  00  00  00  00  00      FE
Select Case colorT
  Case 13000
     value = &H3
  Case 11000
     value = &H3
  Case 9300
     value = &H4
End Select
SendDataBuf(0) = &H55
SendDataBuf(1) = &H16
SendDataBuf(2) = &H1
SendDataBuf(3) = value
SendDataBuf(4) = &H0
SendDataBuf(5) = &H0
SendDataBuf(6) = &H0
SendDataBuf(7) = &H0
SendDataBuf(8) = &H0
SendDataBuf(9) = &H0
SendDataBuf(10) = chksumSend(SendDataBuf)
SendDataBuf(11) = &HFE

Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub SET_USR_R_GAN(USR_R_GAN As Long)
Dim SendDataBuf(0 To 11) As Byte
'55  0A  02  XX  XX  00  00  00  00  00      FE
SendDataBuf(0) = &H55
SendDataBuf(1) = &HA
SendDataBuf(2) = &H2
SendDataBuf(3) = CByte(USR_R_GAN \ 256)
SendDataBuf(4) = CByte(USR_R_GAN Mod 256)
SendDataBuf(5) = &H0
SendDataBuf(6) = &H0
SendDataBuf(7) = &H0
SendDataBuf(8) = &H0
SendDataBuf(9) = &H0
SendDataBuf(10) = chksumSend(SendDataBuf)
SendDataBuf(11) = &HFE
Debug.Print SendDataBuf(10)
Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub SET_USR_G_GAN(USR_G_GAN As Long)
Dim SendDataBuf(0 To 11) As Byte
'55  0B  02  XX  XX  00  00  00  00  00      FE
SendDataBuf(0) = &H55
SendDataBuf(1) = &HB
SendDataBuf(2) = &H2
SendDataBuf(3) = CByte(USR_G_GAN \ 256)
SendDataBuf(4) = CByte(USR_G_GAN Mod 256)
SendDataBuf(5) = &H0
SendDataBuf(6) = &H0
SendDataBuf(7) = &H0
SendDataBuf(8) = &H0
SendDataBuf(9) = &H0
SendDataBuf(10) = chksumSend(SendDataBuf)
SendDataBuf(11) = &HFE

Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub SET_USR_B_GAN(USR_B_GAN As Long)
Dim SendDataBuf(0 To 11) As Byte
'55  0C  02  XX  XX  00  00  00  00  00      FE
SendDataBuf(0) = &H55
SendDataBuf(1) = &HC
SendDataBuf(2) = &H2
SendDataBuf(3) = CByte(USR_B_GAN \ 256)
SendDataBuf(4) = CByte(USR_B_GAN Mod 256)
SendDataBuf(5) = &H0
SendDataBuf(6) = &H0
SendDataBuf(7) = &H0
SendDataBuf(8) = &H0
SendDataBuf(9) = &H0
SendDataBuf(10) = chksumSend(SendDataBuf)
SendDataBuf(11) = &HFE

Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub SET_RGB_GAN(RGB_GAN As REALRGB)
Dim SendDataBuf(0 To 11) As Byte
'55  0A  06  XX  XX  XX  XX  XX  XX  00      FE
SendDataBuf(0) = &H55
SendDataBuf(1) = &HA
SendDataBuf(2) = &H6
SendDataBuf(3) = CByte(RGB_GAN.cRR \ 256)
SendDataBuf(4) = CByte(RGB_GAN.cRR Mod 256)
SendDataBuf(5) = CByte(RGB_GAN.cGG \ 256)
SendDataBuf(6) = CByte(RGB_GAN.cGG Mod 256)
SendDataBuf(7) = CByte(RGB_GAN.cBB \ 256)
SendDataBuf(8) = CByte(RGB_GAN.cBB Mod 256)
SendDataBuf(9) = &H0
SendDataBuf(10) = chksumSend(SendDataBuf)
SendDataBuf(11) = &HFE

Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub SET_Brightness(Brightness As Long)
Dim SendDataBuf(0 To 11) As Byte
'55  37  02  XX  XX  00  00  00  00  00      FE
SendDataBuf(0) = &H55
SendDataBuf(1) = &H37
SendDataBuf(2) = &H2
SendDataBuf(3) = CByte(Brightness \ 256)
SendDataBuf(4) = CByte(Brightness Mod 256)
SendDataBuf(5) = &H0
SendDataBuf(6) = &H0
SendDataBuf(7) = &H0
SendDataBuf(8) = &H0
SendDataBuf(9) = &H0
SendDataBuf(10) = chksumSend(SendDataBuf)
SendDataBuf(11) = &HFE
End Sub

Public Sub SET_Contrast(Contrast As Long)
Dim SendDataBuf(0 To 11) As Byte
'55  39  02  XX  XX  00  00  00  00  00      FE
SendDataBuf(0) = &H55
SendDataBuf(1) = &H39
SendDataBuf(2) = &H2
SendDataBuf(3) = CByte(Contrast \ 256)
SendDataBuf(4) = CByte(Contrast Mod 256)
SendDataBuf(5) = &H0
SendDataBuf(6) = &H0
SendDataBuf(7) = &H0
SendDataBuf(8) = &H0
SendDataBuf(9) = &H0
SendDataBuf(10) = chksumSend(SendDataBuf)
SendDataBuf(11) = &HFE
End Sub

Public Sub SET_COMMAND_RS()
Dim SendDataBuf(0 To 11) As Byte
Dim i As Integer, j As Integer

Form3.Text1 = UCase$(Form3.Text1)
j = 1
 For i = 0 To 11
    If Mid$(Form3.Text1, j, 1) = "" Then Exit For
    If Mid$(Form3.Text1, j, 1) = " " Then j = j + 1
    SendDataBuf(i) = StringToInt(Mid$(Form3.Text1, j, 1)) * 16 + StringToInt(Mid$(Form3.Text1, j + 1, 1))
    j = j + 2
    Debug.Print SendDataBuf(i)
  Next i

SendDataBuf(10) = chksumSend(SendDataBuf)
SendDataBuf(11) = &HFE
Debug.Print SendDataBuf(10)
Form1.MSComm1.Output = SendDataBuf
End Sub

Private Function chksumSend(ByRef data() As Byte) As Byte
Dim i As Integer
chksumSend = &HFF
For i = 1 To 9
  If data(i) = 255 Then
     chksumSend = chksumSend + 1
  Else
     If chksumSend < data(i) Then
     chksumSend = 256 - (data(i) - chksumSend)
     Else
     chksumSend = chksumSend - data(i)
     End If
  End If
Next i
If chksumSend = 255 Then
chksumSend = 0
Else
chksumSend = chksumSend + 1
End If
End Function

Function StringToInt(TS As String) As Byte

Select Case TS
Case Is = "0"
  StringToInt = 0
Case Is = "1"
  StringToInt = 1
Case Is = "2"
  StringToInt = 2
Case Is = "3"
  StringToInt = 3
Case Is = "4"
  StringToInt = 4
Case Is = "5"
  StringToInt = 5
Case Is = "0"
  StringToInt = 0
Case Is = "6"
  StringToInt = 6
Case Is = "7"
  StringToInt = 7
Case Is = "8"
  StringToInt = 8
Case Is = "9"
  StringToInt = 9
Case Is = "A"
  StringToInt = 10
Case Is = "B"
  StringToInt = 11
Case Is = "C"
  StringToInt = 12
Case Is = "D"
  StringToInt = 13
Case Is = "E"
  StringToInt = 14
Case Is = "F"
  StringToInt = 15
Case Is = "P"
  MsgBox ("Command Format is Wrong.")
Case Is = " "
  MsgBox ("Command Format is Wrong.")
End Select

End Function
