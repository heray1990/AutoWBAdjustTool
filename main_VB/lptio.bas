Attribute VB_Name = "LptIo"
Option Explicit


Declare Sub I2c_Start Lib "LPTIO.dll" ()
Declare Sub I2c_RStart Lib "LPTIO.dll" ()
Declare Sub I2c_Stop Lib "LPTIO.dll" ()
Declare Function I2c_Tx Lib "LPTIO.dll" (ByVal portaddr&) As Integer
Declare Function I2c_Rx Lib "LPTIO.dll" (ByRef portaddr As Long, ByVal ack As Integer) As Integer
'Declare Function I2c_ReadString Lib "LPTIO.dll" (ByVal sz As String, ByRef sz1 As String) As Integer
'Declare Function I2c_ReadStringEx Lib "LPTIO.dll" (ByVal str As String) As Long
Declare Function Lpt_Output Lib "LPTIO.dll" (ByVal dat&) As Integer
Declare Function LptioSetDevice Lib "LPTIO.dll" (ByVal device&) As Byte
Declare Function Pll_Clock Lib "LPTIO.dll" (ByVal f As Long, ByVal base_f As Long, ByRef m As Integer, ByRef r As Integer, ByRef n As Integer) As Long
Declare Function Pll_Nandu Lib "LPTIO.dll" (ByVal f As Long, ByVal base_f As Long, ByRef h As Integer, ByRef m As Integer, ByRef l As Integer) As Long
Declare Function Pll_Write Lib "LPTIO.dll" (ByVal dev As Byte, ByVal addr As Byte, ByVal m As Long, ByVal r As Long, ByVal n As Long) As Long
Declare Function uart_connect Lib "LPTIO.dll" (ByVal port As Byte, ByVal baudrate As Long) As Byte
Declare Sub uart_disconnect Lib "LPTIO.dll" ()
Declare Function I2cWriteData Lib "LPTIO.dll" (ByVal device As Byte, ByRef cmd As Byte, ByVal cmdlen As Long, ByRef wdat As Byte, ByVal wdatlen As Long) As Byte 'Boolean
Declare Function I2cReadData Lib "LPTIO.dll" (ByVal device As Byte, ByRef cmd As Byte, ByVal cmdlen As Long, ByRef rdat As Byte, ByVal rdatlen As Long) As Byte 'Boolean
Declare Sub I2cSetClockRate Lib "LPTIO.dll" (ByVal rate As Long)
Declare Sub SetPortVal Lib "LPTIO.dll" (ByVal addr As Byte, ByVal val As Byte)

'Andes ISP
Declare Sub AndesPauseMcu Lib "LPTIO.dll" ()
Declare Sub AndesWriteSpiSr Lib "LPTIO.dll" (ByVal val As Byte)
Declare Sub AndesBurstMode Lib "LPTIO.dll" ()
Declare Sub AndesWriteSpiData Lib "LPTIO.dll" (ByVal addr As Long, ByRef buf As Byte, ByVal leng As Long)
Declare Sub AndesNormalMode Lib "LPTIO.dll" ()
Declare Function AndesReadCrc Lib "LPTIO.dll" (ByVal addr As Long, ByVal leng As Long) As Byte
Declare Sub AndesStartMcu Lib "LPTIO.dll" ()
Declare Function AndesIsp Lib "LPTIO.dll" (ByRef buf As Byte, ByVal leng As Long) As Long
Declare Function CRC8 Lib "LPTIO.dll" (ByVal dat As Byte, ByVal crc As Byte) As Byte

Declare Function GetDeviceNumber Lib "LPTIO.dll" () As Integer

Declare Function I2c_ReadString Lib "LPTIO.dll" (ByRef sz As Byte, ByVal str As String) As Long 'sz: # hex bytes of str, str: read cmd string
Declare Function I2c_WriteString Lib "LPTIO.dll" (ByRef sz As Byte, ByVal str As String) As Long 'sz: # hex bytes of str, str: write cmd string




Global gBank As Byte



' I2C Tool Device Constant
Global Const DEVICE_LPT = 0
Global Const DEVICE_U2C = 1
Global Const DEVICE_WT = 2
Global Const DEVICE_UART = 3
Global Const DEVICE_USBTCON = 4
Global Const DEVICE_FTDI = 8
Global Const DEVICE_AARDVARK = 16
Global DEVICE_USED As Integer


' I2C Protocol Constant
Global Const M2REG_DEVICE_I2C = 0
Global Const M2REG_DEVICE_DDCCI = 1
Global Const M2REG_DEVICE_DDCCI_STANDBY = 2
Global Const M2REG_DEVICE_I2C_INDIRECT = 3
Global Const M2REG_DEVICE_I2C_DIRECT = 4


' ASIC Type Constant
Global Const ASIC_TYPE_MG4 = 0   'Magic-4
Global Const ASIC_TYPE_MG5 = 1   'Magic-5
Global Const ASIC_TYPE_M2I = 2   'M2I
Global Const ASIC_TYPE_M2K = 3   'M2K
Global Const ASIC_TYPE_NUM = 4   'Total number of ASIC type


' ASIC Sub Type Constant
Global Const ASIC_SUB_TYPE_FPGA_VIRTEX_4 = 0
Global Const ASIC_SUB_TYPE_FPGA_VIRTEX_5 = 1
Global Const ASIC_SUB_TYPE_SHUTTLE = 2
Global Const ASIC_SUB_TYPE_ECO_1 = 3
Global Const ASIC_SUB_TYPE_ECO_2 = 4
Global Const ASIC_SUB_TYPE_ECO_3 = 5
Global Const ASIC_SUB_TYPE_MP = 6

