Attribute VB_Name = "Common"
'Option Explicit


' ========================================================
'
'    Common Functions
'
' ========================================================
Public Declare Function GetTickCount Lib "kernel32" () As Long


' ===========================================
'    I2c Control Entry Declaration:
' ===========================================
Global I2C As New I2C



' ============================================
'    Bit definition:
' ============================================
Global Const BIT_31 = &H80000000       'Bit 31
Global Const BIT_30 = &H40000000       'Bit 30
Global Const BIT_29 = &H20000000       'Bit 29
Global Const BIT_28 = &H10000000       'Bit 28
Global Const BIT_27 = &H8000000        'Bit 27
Global Const BIT_26 = &H4000000        'Bit 26
Global Const BIT_25 = &H2000000        'Bit 25
Global Const BIT_24 = &H1000000        'Bit 24

Global Const BIT_23 = &H800000         'Bit 23
Global Const BIT_22 = &H400000         'Bit 22
Global Const BIT_21 = &H200000         'Bit 21
Global Const BIT_20 = &H100000         'Bit 20
Global Const BIT_19 = &H80000          'Bit 19
Global Const BIT_18 = &H40000          'Bit 18
Global Const BIT_17 = &H20000          'Bit 17
Global Const BIT_16 = &H10000          'Bit 16

Global Const BIT_15 = &H8000           'Bit 15
Global Const BIT_14 = &H4000           'Bit 14
Global Const BIT_13 = &H2000           'Bit 13
Global Const BIT_12 = &H1000           'Bit 12
Global Const BIT_11 = &H800            'Bit 11
Global Const BIT_10 = &H400            'Bit 10
Global Const BIT_9 = &H200             'Bit 9
Global Const BIT_8 = &H100             'Bit 8

Global Const BIT_7 = &H80              'Bit 7
Global Const BIT_6 = &H40              'Bit 6
Global Const BIT_5 = &H20              'Bit 5
Global Const BIT_4 = &H10              'Bit 4
Global Const BIT_3 = &H8               'Bit 3
Global Const BIT_2 = &H4               'Bit 2
Global Const BIT_1 = &H2               'Bit 1
Global Const BIT_0 = &H1               'Bit 1


' =======================================
'    Key Definition
' =======================================
Global Const ENTER_KEY = 13
Global Const ESC_KEY = 27


' =======================================
'    4K OR 4K Lite
' =======================================
Global PA8308 As Boolean
Global Port_1 As Boolean







Public Function Double2Long(ByVal val As Double) As Long

If (val >= 4294967296#) Then
  val = -1#
ElseIf (val >= 2147483648#) Then
  val = val - 4294967296#
End If

Double2Long = CLng(val)

End Function


Public Function Long2Double(ByVal val As Long) As Double

If (val < 0) Then
  Long2Double = CDbl(4294967296# + val)
Else
  Long2Double = CDbl(val)
End If

End Function


Public Function Int2HexStr(src As Variant, ByRef dst As String, l As Byte)

Dim ss As String
ss = Hex(src)
PatchString ss, l
dst = ss

End Function


Public Function Int2HexStr2(src As Variant, l As Byte) As Variant

Dim ss As String
ss = Hex(src)
PatchString ss, l
Int2HexStr2 = ss

End Function


Public Function Int2Str(ByVal src As Variant, ByVal l As Byte) As Variant

Dim ss As String
ss = CStr(src)
PatchString ss, l
Int2Str = ss

End Function


Public Function HexStr2Int(ByVal src As Variant, ByVal l As Byte)

Dim i As Integer
Dim ss As String
Dim res As Variant
Dim res0 As Variant

ss = src
PatchString ss, l
res = 0
For i = 0 To l / 2 - 1
    res0 = val("&h" & Mid(ss, l - 2 * i - 1, 2))
    res = res + res0 * 256 ^ i
Next
'HexStr2Int = val("&h" & ss)
HexStr2Int = res

End Function


Public Sub PatchString(ByRef i As String, ByVal j As Long)

Dim a As Integer
Dim b As Integer
a = Len(i)
If (a <= j) Then
For b = a To j - 1
    i = "0" & i
Next b
Else
i = Right(i, j)
End If

End Sub


Public Function sPatchString(ByRef i As String, ByVal j As Long) As String

Dim a As Integer
Dim b As Integer
a = Len(i)
If (a <= j) Then
  For b = a To j - 1
    i = "0" & i
  Next b
Else
  i = Right(i, j)
End If
sPatchString = i

End Function


Public Function BitClear(n As Variant, mask As Variant) As Variant

ret = n And Not (mask)
BitClear = ret

End Function


Public Function BitSet(n As Variant, mask As Variant) As Variant

ret = n Or mask
BitSet = ret

End Function


Public Function BitInvert(n As Variant, mask As Variant) As Variant

ret = n Xor mask
BitInvert = ret

End Function


Public Function isBitSet(n As Variant, mask As Variant) As Variant
'    MsgBox "n=" & Hex(n) & " and mask=" & Hex(mask) & " = " & Hex(n And mask)
If (n And mask) <> 0 Then
    isBitSet = 1
Else
    isBitSet = 0
End If

End Function


Public Function isValidateKey(KeyCode As Byte, ByRef key As Byte) As Boolean

Dim success As Boolean

Select Case KeyCode
    Case 48 To 57
        key = KeyCode - 48
        success = True
    Case 65 To 71
        key = KeyCode - 55
        success = True
    Case Else
        success = False
End Select
isValidateKey = success

End Function


Public Function Bool2Val(d As Boolean) As Byte

If (d = True) Then
    Bool2Val = 1
Else
    Bool2Val = 0
End If
    
End Function


Public Function Val2Bool(d) As Boolean

If (d <> 0) Then
    Val2Bool = True
Else
    Val2Bool = False
End If

End Function


Public Function bit(ByVal m As Long, ByVal n As Long) As Boolean

If (m And 2 ^ n) Then
    bit = True
Else
    bit = False
End If

End Function



Public Function bits(ByVal val As Long, ByVal msb As Long, ByVal lsb As Long) As Long
    Dim signbit As Long
    If (val And &H80000000) <> 0 Then signbit = 1 Else signbit = 0
    If lsb <> 0 Then
        val = val And &H7FFFFFFF
        val = Int(val / (2 ^ lsb))
        If signbit = 1 Then val = val + &H40000000 / (2 ^ (lsb - 1))
    End If
    msb = msb - lsb
    If msb = 31 Then
        bits = val
    Else
        mask = 2 ^ (msb + 1) - 1
        val = val And mask
        bits = val
    End If
End Function


Public Function EnBit(ByVal m As Long, ByVal n As Long) As Long

m = m Or 2 ^ n
EnBit = m

End Function


Public Function DisBit(ByVal m As Long, ByVal n As Long) As Long

m = m And (&HFFFFFFFF - 2 ^ n)
DisBit = m

End Function


Public Function TimeDelay(ByVal TT As Long)

 Dim T As Long    ' 声明一个长整数, 记录计数值
 T = GetTickCount()   ' 取得系统计数值
 Do
   DoEvents
   If GetTickCount - T < 0 Then T = GetTickCount ' 归零
 Loop Until GetTickCount - T >= TT  ' 计数延迟是否达到
' Print "delay OK"

End Function





Public Function SetBit8(ByRef src As Byte, ByVal str As Byte, ByVal dat As Byte) As Byte

SetBit8 = SetBits8(src, str, 1, dat)

End Function


Public Function SetBits8(ByRef src As Byte, ByVal str As Byte, ByVal lens As Byte, ByVal dat As Byte) As Byte

Dim mask As Byte
Dim Shift As Byte

If (str + lens) > 8 Then
  SetBits8 = 0
  MsgBox "Error: (str+lens) > 8"
  Exit Function
End If

mask = (2 ^ lens) - 1
Shift = (2 ^ str)
dat = (dat And mask) * Shift
mask = mask * (2 ^ str)

src = src And (Not mask) Or dat
SetBits8 = src

End Function


'
' Bank-List Handler Functions
'
Public Function ComboLoadBanks(ByRef cmb As ComboBox, ByVal has_all As Boolean) As Integer
    cmb.Clear
    cmb.AddItem "00: SYS_BANK", 0
    cmb.AddItem "01: MISC_BANK", 1
    cmb.AddItem "02: VBORX0_BANK", 2
    cmb.AddItem "03: VBORX1_BANK", 3
    cmb.AddItem "04: INP_BANK", 4
    cmb.AddItem "05: MLV_BANK", 5
    cmb.AddItem "06: DDR3_MC_BANK", 6
    cmb.AddItem "07: DAC_IF_BANK", 7
    cmb.AddItem "08: PB2AXI_BANK", 8
    cmb.AddItem "09: REGUSER_BANK", 9
    cmb.AddItem "10: Core2", 10
    cmb.AddItem "11: Core2", 11
    cmb.AddItem "30: Core1 Top", 12
    cmb.AddItem "34: Daphne", 13
    cmb.AddItem "35: Decont LUT", 14
    cmb.AddItem "36: Decont", 15
    cmb.AddItem "37: Core1 37", 16
    cmb.AddItem "38: Core1 38", 17
    cmb.AddItem "39: Core1 39", 18
    cmb.AddItem "3A: QuattronPlus", 19
    If (has_all) Then cmb.AddItem "All"
End Function


Public Function Index2Bank(ByVal idx As Integer) As Integer
  
  'Dim RefCombo As ComboBox
  'RefCombo = Form1.ComboBankRef

'  If (idx = (Form1.ComboBankRef.ListCount - 1)) Then
'    Index2Bank = &HFF
'  Else
'    'RefCombo.ListIndex = idx
'    'Index2Bank = CLng(val("&H" & Mid$(RefCombo.Text, 1, 2))) And &HFF
'    Index2Bank = CLng(val("&H" & Mid$(Form1.ComboBankRef.List(idx), 1, 2) & "#"))
'  End If
  
'  Select Case (idx)
'    Case 0
'      Index2Bank = SYS_BANK
'    Case 1
'      Index2Bank = MISC_BANK
'    Case 2
'      Index2Bank = VBORX0_BANK
'    Case 3
'      Index2Bank = VBORX1_BANK
'    Case 4
'      Index2Bank = INP_BANK
'    Case 5
'      Index2Bank = MLV_BANK
'    Case 6
'      Index2Bank = DDR3_MC_BANK
'    Case 7
'      Index2Bank = DAC_IF_BANK
'    Case 8
'      Index2Bank = PB2AXI_BANK
'    Case 9
'      Index2Bank = REGUSER_BANK
'    Case 10
'      Index2Bank = SHARP_CORE2_10
'    Case 11
'      Index2Bank = SHARP_CORE2_11
'    Case 12
'      Index2Bank = SHARP_CORE1_TOP
'    Case 13
'      Index2Bank = SHARP_DAPHNE
'    Case 14
'      Index2Bank = SHARP_DECODE
'    Case 15
'      Index2Bank = SHARP_QUATTRONPLUS
'    Case 16
'      Index2Bank = SHARP_35
'    Case 17
'      Index2Bank = SHARP_37
'    Case 18
'      Index2Bank = SHARP_38
'    Case 19
'      Index2Bank = SHARP_39
'    Case Else
'      Index2Bank = SYS_BANK
'  End Select
End Function



Public Function Bank2Index(ByVal bank As Byte) As Integer
  Dim idx As Integer
  Dim getBank As Byte
  'Dim RefCombo As ComboBox
  'RefCombo = Form1.ComboBankRef
  
'  For idx = 0 To (Form1.ComboBankRef.ListCount - 2)
'    'RefCombo.ListIndex = idx
'    'getBank = CLng(val("&H" & Mid$(RefCombo.Text, 1, 2))) And &HFF
'    getBank = CLng(val("&H" & Mid$(Form1.ComboBankRef.List(idx), 1, 2) & "#"))
'    If (bank = getBank) Then
'      Bank2Index = idx
'      Exit Function
'    End If
'  Next idx
'
'  Bank2Index = -1
'
'  Select Case (bank)
'    Case SYS_BANK
'      Bank2Index = 0
'    Case MISC_BANK
'      Bank2Index = 1
'    Case VBORX0_BANK
'      Bank2Index = 2
'    Case VBORX1_BANK
'      Bank2Index = 3
'    Case INP_BANK
'      Bank2Index = 4
'    Case MLV_BANK
'      Bank2Index = 5
'    Case DDR3_MC_BANK
'      Bank2Index = 6
'    Case DAC_IF_BANK
'      Bank2Index = 7
'    Case PB2AXI_BANK
'      Bank2Index = 8
'    Case REGUSER_BANK
'      Bank2Index = 9
'    Case SHARP_CORE2_10
'      Bank2Index = 10
'    Case SHARP_CORE2_11
'      Bank2Index = 11
'    Case SHARP_CORE1_TOP
'      Bank2Index = 12
'    Case SHARP_DAPHNE
'      Bank2Index = 13
'    Case SHARP_DECODE
'      Bank2Index = 14
'    Case SHARP_QUATTRONPLUS
'      Bank2Index = 15
'    Case SHARP_35
'      Bank2Index = 16
'    Case SHARP_37
'      Bank2Index = 17
'    Case SHARP_38
'      Bank2Index = 18
'    Case SHARP_39
'      Bank2Index = 19
'    Case Else
'      Bank2Index = -1
'      'MsgBox "Unknown Bank: " & Hex$(bank)
'  End Select
End Function



