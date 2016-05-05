VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "I2C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public device As Integer
'Public aardvark_handle As Long

Public I2cDebug As Boolean


Public Function Rd( _
  ByVal devID As Byte, _
  ByRef rcmd() As Byte, _
  ByVal rcmdlen As Long, _
  ByRef rdat() As Byte, _
  ByVal rdatlen As Long _
) As Boolean

Dim ack As Boolean

ack = False

If (I2cReadData(devID, rcmd(0), rcmdlen, rdat(0), rdatlen) = 1) Then
    ack = True
End If

'Select Case (device)
'    Case DEVICE_FTDI, DEVICE_USBTCON
        ' Himax FTDI or Cypress Tool
'        If (I2cReadData(devID, rcmd(0), rcmdlen, rdat(0), rdatlen) = 1) Then
'           ack = True
'        End If

'    Case DEVICE_AARDVARK
        ' SHARP AardVark Tool
       'devID = devID \ 2
       'If (aa_i2c_write(aardvark_handle, devID, AA_I2C_NO_FLAGS, rcmdlen, rcmd(0)) > 0) Then
       '    If (aa_i2c_read(aardvark_handle, devID, AA_I2C_NO_FLAGS, rdatlen, rdat(0)) > 0) Then
       '        ack = True
       '    End If
       'End If

'    Case Else
        ' Un-Registered I2C Device
'        PrintLog "Invalid I2C Device (" & CStr(device) & ")"

'End Select

Rd = ack

End Function



Public Function Wr( _
  ByVal devID As Byte, _
  ByRef wdat() As Byte, _
  ByVal wlen As Long _
) As Boolean

Dim ack As Boolean

ack = False

Select Case (device)
    Case DEVICE_FTDI, DEVICE_USBTCON
        ' Himax FTDI or Cypress Tool
        If (I2cWriteData(devID, wdat(0), 0, wdat(0), wlen) = 1) Then
            ack = True
        End If
  
    Case DEVICE_AARDVARK
        ' SHARP AardVark Tool
       'devID = devID \ 2
       'If (aa_i2c_write(aardvark_handle, devID, AA_I2C_NO_FLAGS, wlen, wdat(0)) > 0) Then
       '    ack = True
       'End If

    Case Else
        ' Un-Registered I2C Device
        'PrintLog "Invalid I2C Device (" & CStr(device) & ")"

End Select

Wr = ack

End Function



Public Sub SetClockRateKHz(ByVal val As Long)
Call I2cSetClockRate(val)
'Select Case (device)
'    Case DEVICE_FTDI, DEVICE_USBTCON
        ' Himax FTDI or Cypress Tool
'        Call I2cSetClockRate(val)
  
'    Case DEVICE_AARDVARK
        ' SHARP AardVark Tool

'    Case Else
        ' Un-Registered I2C Device
'        PrintLog "Invalid I2C Device (" & CStr(device) & ")"

'End Select

End Sub




