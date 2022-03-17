VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnsiFormatRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IFormatRules

Private Const mTextQuote As String = "'"
Private Const mDoubleQuote As String = mTextQuote & mTextQuote
Private Const mLeftQuote As String = """"
Private Const mRightQuote As String = """"

Private mDbTimeStampOffsetSize As Long
Private mDbTime2Size As Long
Private mPointerSize As Long
Private mPointer As LongPtr
Private mDbTimeStampOffset As DBTIMESTAMPOFFSET
Private mDbTime2 As DBTime2

Private mBinaryConverter As IBinaryConverter

Private Sub Class_Initialize()
  mDbTimeStampOffsetSize = LenB(mDbTimeStampOffset)
  mDbTime2Size = LenB(mDbTime2)
  mPointerSize = LenB(mPointer)
End Sub


Private Sub IFormatRules_AppendFormattedBinary(targetString As IStringBuilder, bin As Variant, Optional Profile As BinaryProfile = Nothing)
  targetString.Add IFormatRules_GetFormattedBinary(bin, Profile:=Profile)
End Sub

Private Sub IFormatRules_AppendFormattedBoolean(targetString As IStringBuilder, bool As Variant)
  targetString.Add IFormatRules_GetFormattedBoolean(bool)
End Sub

Private Sub IFormatRules_AppendFormattedDate(targetString As IStringBuilder, dte As Variant)
  targetString.Add IFormatRules_GetFormattedDate(dte)
End Sub

Private Sub IFormatRules_AppendFormattedDbDateTimeOffset(targetString As IStringBuilder, dte As Variant)
targetString.Add IFormatRules_GetFormattedDbDateTimeOffset(dte)
End Sub

Private Sub IFormatRules_AppendFormattedDbTime(targetString As IStringBuilder, tme As Variant)
  targetString.Add IFormatRules_GetFormattedDbTime(tme)
End Sub

Private Sub IFormatRules_AppendFormattedNullableNumber(targetString As IStringBuilder, num As Variant)
    targetString.Add IFormatRules_GetFormattedNullableNumber(num)
End Sub

Private Sub IFormatRules_AppendFormattedText(targetString As IStringBuilder, txt As Variant)
  targetString.Add IFormatRules_GetFormattedText(txt)
End Sub

Private Function IFormatRules_GetFormattedDbDateTimeOffset(dte As Variant) As String
  Dim b() As Byte
  Const DIGIT As LongLong = &H100^
  Const DIGIT2 As LongLong = &H10000^
  Const DIGIT3 As LongLong = &H1000000^
  
  If VarType(dte) = (vbArray Or vbByte) Then
    'With mDbTimeStampOffset
      b = dte
      ' This is ugly, but it's WAY faster than using RtlMoveMemory to load the data to the type and doing it that way.
      IFormatRules_GetFormattedDbDateTimeOffset = "'" & (b(0) + DIGIT * b(1)) & "-" & (b(2)) & "-" & (b(4)) & " " & (b(6)) & ":" & (b(8)) & ":" & (b(10)) & "." & _
        Right$("0000000" & ((b(12) + b(13) * DIGIT + b(14) * DIGIT2 + b(15) * DIGIT3) / 100), 7) & " " & Switch(b(16) >= 0, "+", 1 = 1, "-") & (b(16)) & ":" & (b(18)) & "'"
'      IFormatRules_GetFormattedDbDateTimeOffset = "'" & .year & "-" & .month & "-" & .day & " " & .hour & ":" & .minute & ":" & .second & "." & _
        Right$("0000000" & (.fraction / 100), 7) & " " & Switch(.timezone_hour >= 0, "+", 1 = 1, "-") & .timezone_hour & ":" & .timezone_minute & "'"
    'End With
  Else
    If IsDate(dte) Then
      IFormatRules_GetFormattedDbDateTimeOffset = IFormatRules_GetFormattedDate(dte)
      Exit Function
    End If
    IFormatRules_GetFormattedDbDateTimeOffset = "NULL"
    Err.Raise 516, "AnsiFormatRules.GetFormattedDbTime", "Invalid time format."
  End If
End Function

Private Function IFormatRules_GetFormattedBinary(bin As Variant, Optional Profile As BinaryProfile = Nothing) As String

  Call mBinaryConverter.ToBinary(IFormatRules_GetFormattedBinary, bin)
End Function

Private Function IFormatRules_GetFormattedBoolean(txt As Variant) As String
  If IsNull(txt) Then
    IFormatRules_GetFormattedBoolean = "NULL"
  Else
    IFormatRules_GetFormattedBoolean = Switch(txt = True, "1", txt = False, "0")
  End If
End Function

Private Function IFormatRules_GetFormattedDbTime(tme As Variant) As String
  Dim b() As Byte
  If VarType(tme) = (vbArray Or vbByte) Then
    b = tme
    'With mDbTime2
      ' This is ugly, but it's WAY faster than using RtlMoveMemory to load the data to the type and doing it that way.
      IFormatRules_GetFormattedDbTime = "'" & b(0) & ":" & b(2) & ":" & b(4) & "." & Format("0000000", (b(8) + b(9) * &H100^ + b(10) * &H10000^ + b(11) * &H1000000^) / 100) & "'"
      'IFormatRules_GetFormattedDbTime = "'" & .hour & ":" & .minute & ":" & .second & "." & Format("0000000", .fraction / 100) & "'"
    'End With
  Else
    IFormatRules_GetFormattedDbTime = "NULL"
    Err.Raise 516, "AnsiFormatRules.GetFormattedDbTime", "Invalid time format."
  End If

End Function

Private Function IFormatRules_GetFormattedNullableNumber(num As Variant) As String
  If IsNull(num) Then
    IFormatRules_GetFormattedNullableNumber = "NULL"
  Else
    IFormatRules_GetFormattedNullableNumber = num
  End If
End Function

Private Property Get IFormatRules_LeftQuote() As String
  IFormatRules_LeftQuote = mLeftQuote
End Property

Private Function IFormatRules_QuoteName(name As String) As String
  IFormatRules_QuoteName = IFormatRules_LeftQuote & Replace(name, IFormatRules_RightQuote, VBA.String$(2, IFormatRules_RightQuote)) & IFormatRules_RightQuote
End Function

Private Property Get IFormatRules_RightQuote() As String
  IFormatRules_RightQuote = mRightQuote
End Property

Private Function IFormatRules_GetFormattedText(txt As Variant) As String
  If IsNull(txt) Then
    IFormatRules_GetFormattedText = "NULL"
  Else
    IFormatRules_GetFormattedText = mTextQuote & Replace(txt, mTextQuote, mDoubleQuote) & mTextQuote
  End If
End Function

Private Property Get IFormatRules_TextDelimiter() As String
  IFormatRules_TextDelimiter = mTextQuote
End Property

Private Function IFormatRules_GetFormattedDate(dte As Variant) As String
  Dim f As ADODB.field
  Dim v As Variant
  Dim formatString As String
  Dim DataType As ADODB.DataTypeEnum
  
  Const adDbDateFormat As String = "'YYYY-MM-DD'"
  Const defaultFormat As String = "'YYYY-MM-DD HH:NN:SS'"
  
  If TypeOf dte Is ADODB.field Then
    Set f = dte
    DataType = f.Type
    v = f.value
  Else
    v = dte
    If IsDate(v) Then
      If (CDate(v) - Int(CDate(v))) = 0 Then
      ' The extra work on the client side is _
        probably worth the network savings...?
        DataType = adDBDate
      Else
      ' Use most detailed format
        DataType = adDBTimeStamp
      End If
    Else
    ' Assume custom date formatting
      DataType = adVarWChar
    End If
  End If
  
  If IsNull(v) Then
    IFormatRules_GetFormattedDate = "NULL"
    Exit Function
  End If
  
  Select Case DataType
    Case adDBDate
      formatString = adDbDateFormat
    Case ADODB.DataTypeEnum.adDBDate, ADODB.DataTypeEnum.adDBTime, ADODB.DataTypeEnum.adDBTimeStamp, ADODB.DataTypeEnum.adFileTime
      formatString = defaultFormat
    Case Else
      If IsDate(v) Then
        IFormatRules_GetFormattedDate = mTextQuote & v & mTextQuote
        Exit Function
      Else
      ' If you're passing through a different date format it's on you to ensure that it's valid. _
        This is one way to get around the "OLEDB trims milliseconds" problem.
        IFormatRules_GetFormattedDate = mTextQuote & Replace(v, mTextQuote, mDoubleQuote) & mTextQuote
        Exit Function
      End If
  End Select
  IFormatRules_GetFormattedDate = Format(dte, formatString)
End Function
