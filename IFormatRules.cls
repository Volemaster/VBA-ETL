VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormatRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Property Get LeftQuote() As String

End Property

Public Property Get RightQuote() As String

End Property

Public Function QuoteName(name As String) As String

End Function

Public Function GetFormattedDate(dte As Variant) As String
  
End Function

Public Function GetFormattedText(txt As Variant) As String

End Function

Public Function GetFormattedNullableNumber(txt As Variant) As String

End Function

Public Function GetFormattedBinary(bin As Variant, Optional Profile As BinaryProfile = Nothing) As String

End Function

Public Function GetFormattedBoolean(txt As Variant) As String

End Function

Public Function GetFormattedDbTime(tme As Variant) As String

End Function

Public Function GetFormattedDbDateTimeOffset(dte As Variant) As String

End Function

Public Property Get TextDelimiter() As String

End Property

Public Sub AppendFormattedDate(targetString As IStringBuilder, dte As Variant)

End Sub

Public Sub AppendFormattedText(targetString As IStringBuilder, txt As Variant)

End Sub

Public Sub AppendFormattedNullableNumber(targetString As IStringBuilder, num As Variant)

End Sub

Public Sub AppendFormattedBinary(targetString As IStringBuilder, bin As Variant, Optional Profile As BinaryProfile = Nothing)

End Sub

Public Sub AppendFormattedBoolean(targetString As IStringBuilder, bool As Variant)

End Sub

Public Sub AppendFormattedDbTime(targetString As IStringBuilder, tme As Variant)

End Sub

Public Sub AppendFormattedDbDateTimeOffset(targetString As IStringBuilder, dte As Variant)

End Sub

