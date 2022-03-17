VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mFieldName As String
Private mNativeType As String
Private mAdoField As ADODB.field
Private mNullable As Boolean
Private mFormatRule As FormatRuleEnum
Private mBinaryType As ADODB.DataTypeEnum
Private mBinaryProfile As BinaryProfile

Implements IFieldInfo

Private Sub Class_Initialize()
  mBinaryType = adBinary
End Sub

Private Function IFieldInfo_AppendValue(formatter As IFormatRules, target As IStringBuilder) As Variant
  Dim value As Variant
  value = mAdoField.value
  
  Select Case mFormatRule
    Case FormatRuleEnum.frNone
      target.Add CStr(value)
    Case FormatRuleEnum.frNullOnly
      formatter.AppendFormattedNullableNumber target, value
    Case FormatRuleEnum.frText
      formatter.AppendFormattedText target, value
    Case FormatRuleEnum.frDate
      formatter.AppendFormattedDate target, value
    Case FormatRuleEnum.frBinary
      formatter.AppendFormattedBinary target, value, mBinaryProfile
    Case FormatRuleEnum.frBoolean
      formatter.AppendFormattedBoolean target, value
    Case FormatRuleEnum.frDbTime
      formatter.AppendFormattedDbTime target, value
    Case FormatRuleEnum.frDbDateTimeOffset
      formatter.AppendFormattedDbDateTimeOffset target, value
  End Select
End Function

Private Property Let IFieldInfo_BinaryType(RHS As ADODB.DataTypeEnum)
  mBinaryType = RHS
End Property

Private Property Get IFieldInfo_BinaryType() As ADODB.DataTypeEnum
  IFieldInfo_BinaryType = mBinaryType
End Property

Public Property Get BinaryProfile() As BinaryProfile
  Set BinaryProfile = mBinaryProfile
End Property

Private Property Get IFieldInfo_FieldName() As String
  IFieldInfo_FieldName = mFieldName
End Property

Private Function IFieldInfo_FormatValue(formatter As IFormatRules) As String
  Dim value As Variant
  value = mAdoField.value
  Select Case mFormatRule
    Case FormatRuleEnum.frNone
      IFieldInfo_FormatValue = CStr(value)
    Case FormatRuleEnum.frNullOnly
      IFieldInfo_FormatValue = formatter.GetFormattedNullableNumber(value)
    Case FormatRuleEnum.frText
      IFieldInfo_FormatValue = formatter.GetFormattedText(value)
    Case FormatRuleEnum.frDate
      IFieldInfo_FormatValue = formatter.GetFormattedDate(value)
    Case FormatRuleEnum.frBinary
      IFieldInfo_FormatValue = formatter.GetFormattedBinary(value)
    Case FormatRuleEnum.frBoolean
      IFieldInfo_FormatValue = formatter.GetFormattedBoolean(value)
    Case FormatRuleEnum.frDbTime
      IFieldInfo_FormatValue = formatter.GetFormattedDbTime(value)
    Case FormatRuleEnum.frDbDateTimeOffset
      IFieldInfo_FormatValue = formatter.GetFormattedDbDateTimeOffset(value)
  End Select
End Function

Private Property Get IFieldInfo_NativeType() As String
  IFieldInfo_NativeType = mNativeType
End Property

Private Property Get IFieldInfo_AdoField() As ADODB.field
  Set IFieldInfo_AdoField = mAdoField
End Property

Public Sub Setup(FieldName As String, NativeType As String, Nullable As Boolean, AdoField As ADODB.field, Optional ExplicitDataType As ADODB.DataTypeEnum = ADODB.DataTypeEnum.adBinary)
  mFieldName = FieldName
  mNativeType = NativeType
  Set mAdoField = AdoField
  mNullable = Nullable
  Dim scale_ As Variant
  If mAdoField Is Nothing Then
  ' Text is the most conservative formatting... also the slowest.
    mFormatRule = frText
    Exit Sub
  End If
  If ExplicitDataType = adBinary Then
    ExplicitDataType = mAdoField.Type
  End If
  Select Case ExplicitDataType
    Case ADODB.DataTypeEnum.adBoolean
      mFormatRule = frBoolean
    Case ADODB.DataTypeEnum.adBinary, _
          ADODB.DataTypeEnum.adLongVarBinary, _
          ADODB.DataTypeEnum.adVarBinary
      mFormatRule = frBinary
    Case ADODB.DataTypeEnum.adBigInt, _
          ADODB.DataTypeEnum.adDecimal, _
          ADODB.DataTypeEnum.adDouble, _
          ADODB.DataTypeEnum.adInteger, _
          ADODB.DataTypeEnum.adNumeric, _
          ADODB.DataTypeEnum.adSmallInt, _
          ADODB.DataTypeEnum.adSingle, _
          ADODB.DataTypeEnum.adTinyInt, _
          ADODB.DataTypeEnum.adUnsignedBigInt, _
          ADODB.DataTypeEnum.adUnsignedInt, _
          ADODB.DataTypeEnum.adUnsignedSmallInt, _
          ADODB.DataTypeEnum.adUnsignedTinyInt, _
          ADODB.DataTypeEnum.adVarNumeric
        mFormatRule = frNullOnly
      Case ADODB.DataTypeEnum.adBSTR
        mFormatRule = frText
      Case ADODB.DataTypeEnum.adChapter
        Err.Raise 513, "Ado->Sql", "Invalid ADODB.DataTypeEnum value"
      Case ADODB.DataTypeEnum.adChar
        mFormatRule = frText
      Case ADODB.DataTypeEnum.adCurrency
        mFormatRule = frNullOnly
      Case ADODB.DataTypeEnum.adDate
        mFormatRule = frDate
      Case ADODB.DataTypeEnum.adDBDate
        mFormatRule = frDate
      Case ADODB.DataTypeEnum.adDBTime
        mFormatRule = frDbTime
      Case ADODB.DataTypeEnum.adDBTimeStamp
        mFormatRule = frDbDateTimeOffset
      Case ADODB.DataTypeEnum.adEmpty
        Err.Raise 513, "Ado->Sql", "Invalid ADODB.DataTypeEnum value"
      Case ADODB.DataTypeEnum.adError
        Err.Raise 513, "Ado->Sql", "Invalid ADODB.DataTypeEnum value"
      Case ADODB.DataTypeEnum.adFileTime
        mFormatRule = frDate
      Case ADODB.DataTypeEnum.adGUID
        mFormatRule = frText
      Case ADODB.DataTypeEnum.adLongVarChar
        mFormatRule = frText
      Case ADODB.DataTypeEnum.adLongVarWChar
        mFormatRule = frText
      Case ADODB.DataTypeEnum.adPropVariant
        mFormatRule = frText
      
      Case ADODB.DataTypeEnum.adUserDefined
        Err.Raise 513, "Ado->Sql", "Invalid ADODB.DataTypeEnum value"
      
      Case ADODB.DataTypeEnum.adVarChar
        mFormatRule = frText
      Case ADODB.DataTypeEnum.adVariant
        mFormatRule = frText
      Case ADODB.DataTypeEnum.adVarWChar
        mFormatRule = frText
      Case ADODB.DataTypeEnum.adWChar
        mFormatRule = frText
    End Select
    
    
  
End Sub

Private Property Get IFieldInfo_Nullable() As Boolean
  IFieldInfo_Nullable = mNullable
End Property

Private Function IFieldInfo_ToString(formatter As IFormatRules, Optional formatOption As FieldInfoOptionsEnum = FieldInfoOptionsEnum.fioNameOnly) As String
  Select Case formatOption
    Case FieldInfoOptionsEnum.fioDeclaration
      IFieldInfo_ToString = formatter.QuoteName(mFieldName) & " " & mNativeType & " " & Switch(mNullable = True, "NULL", False, "NOT NULL")
    Case FieldInfoOptionsEnum.fioNameOnly
      IFieldInfo_ToString = formatter.QuoteName(mFieldName)
    Case Else
      Err.Raise 5, "FieldInfo.ToString", "Invalid formatOption specified."
  End Select
End Function
