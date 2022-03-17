VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlFormatRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mTextQuote As String = "'"
Private Const mDoubleQuote As String = mTextQuote & mTextQuote

Implements IFormatRules

Private mAnsiFormatter As IFormatRules
Private mBinaryConverter As IBinaryConverter

Private Sub Class_Initialize()
  Set mAnsiFormatter = New AnsiFormatRules
  Set mBinaryConverter = New BinaryConverter
End Sub

Private Sub IFormatRules_AppendFormattedBinary(targetString As IStringBuilder, bin As Variant, Optional Profile As BinaryProfile = Nothing)
  mBinaryConverter.AppendBinary targetString, bin, Profile:=Profile
End Sub

Private Sub IFormatRules_AppendFormattedBoolean(targetString As IStringBuilder, bool As Variant)
  targetString.Add IFormatRules_GetFormattedBoolean(bool)
End Sub

Private Sub IFormatRules_AppendFormattedDate(targetString As IStringBuilder, dte As Variant)
  targetString.Add IFormatRules_GetFormattedDate(dte)
End Sub

Private Sub IFormatRules_AppendFormattedDbDateTimeOffset(targetString As IStringBuilder, dte As Variant)
  mAnsiFormatter.AppendFormattedDbDateTimeOffset targetString, dte
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

Private Function IFormatRules_GetDateFromDBTimeStamp(value As Variant) As Byte()

End Function

Private Function IFormatRules_GetFormattedBinary(bin As Variant, Optional Profile As BinaryProfile = Nothing) As String
  Call mBinaryConverter.ToBinary(IFormatRules_GetFormattedBinary, bin, Profile:=Profile)
End Function

Private Function IFormatRules_GetFormattedBoolean(txt As Variant) As String
If IsNull(txt) Then
  IFormatRules_GetFormattedBoolean = "NULL"
Else
  If CBool(txt) Then
    IFormatRules_GetFormattedBoolean = "1"
  Else
    IFormatRules_GetFormattedBoolean = "0"
  End If
End If
End Function

Private Function IFormatRules_GetFormattedDate(dte As Variant) As String
  IFormatRules_GetFormattedDate = mAnsiFormatter.GetFormattedDate(dte)
End Function

Private Function IFormatRules_GetFormattedDbDateTimeOffset(dte As Variant) As String
  IFormatRules_GetFormattedDbDateTimeOffset = mAnsiFormatter.GetFormattedDbDateTimeOffset(dte)
End Function

Private Function IFormatRules_GetFormattedDbTime(tme As Variant) As String
  IFormatRules_GetFormattedDbTime = mAnsiFormatter.GetFormattedDbTime(tme)
End Function

Private Function IFormatRules_GetFormattedNullableNumber(txt As Variant) As String
  IFormatRules_GetFormattedNullableNumber = mAnsiFormatter.GetFormattedNullableNumber(txt)
End Function

Private Function IFormatRules_GetFormattedText(txt As Variant) As String
  IFormatRules_GetFormattedText = mAnsiFormatter.GetFormattedText(txt)
End Function

Private Function IFormatRules_GetTimeFromDbTime(value As Variant) As Byte()
  Dim oletype As DBTime2
  Dim b(0 To 5) As Byte
  Dim millionths As LongLong
  Const frac As LongLong = 10000000
  
  If VarType(value) And (vbArray Or vbByte) = (vbArray Or vbByte) Then
    With oletype
      .hour = value(0)
      .minute = value(2)
      .second = value(4)
      'value(6) and value(7) are padding
      .fraction = value(8) + value(9) * &H10& + value(10) * &H100& + value(11) * &H1000&
    End With
  End If
  With oletype
    millionths = _
      .fraction + _
      .second * frac + _
      .minute * frac * 60 + _
      .hour * frac * 60 * 60
  End With
  CopyMem VarPtr(b(1)), VarPtr(millionths), 5
  b(0) = 7
  IFormatRules_GetTimeFromDbTime = b
  
End Function

Private Property Get IFormatRules_TextDelimiter() As String
  IFormatRules_TextDelimiter = mAnsiFormatter.TextDelimiter
End Property

' These are called infrequently enough that we'll go through the interface
Private Property Get IFormatRules_LeftQuote() As String
  IFormatRules_LeftQuote = "["
End Property

Private Function IFormatRules_QuoteName(name As String) As String
  IFormatRules_QuoteName = IFormatRules_LeftQuote & Replace(name, IFormatRules_RightQuote, VBA.String$(2, IFormatRules_RightQuote)) & IFormatRules_RightQuote
End Function

Private Property Get IFormatRules_RightQuote() As String
  IFormatRules_RightQuote = "]"
End Property
