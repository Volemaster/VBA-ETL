VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlFieldConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IFieldConverter
Implements IFieldConverterSetup

Private mFormatter As IFormatRules
Private mGuessBinaryTypes As Boolean

Private Sub Class_Initialize()
  mGuessBinaryTypes = True
End Sub

Private Property Get IFieldConverter_Formatter() As IFormatRules
  Set IFieldConverter_Formatter = mFormatter
End Property

Private Function IFieldConverter_GetDeclaration(FieldInfo As IFieldInfo) As String
  Dim nullType As String
  With FieldInfo
    If .Nullable Then
      nullType = "NULL"
    Else
      nullType = "NOT NULL"
    End If
    IFieldConverter_GetDeclaration = cFormat("%s %s %s", QuoteName(.FieldName), .NativeType, nullType)
  End With
End Function

Private Function QuoteName(name As String) As String
  QuoteName = mFormatter.QuoteName(name)
End Function

Private Function IFieldConverter_GetFieldInfo(fld As ADODB.field) As IFieldInfo
  Dim sqlType As String, forceSize As Variant
  Dim scale_ As Variant, precision_ As Variant, size_ As Variant
  Dim nullType As String
  Dim fullType As String
  Dim fi As FieldInfo
  Dim Nullable As Boolean
  Dim value As Variant
  Dim functionalTypes As ADODB.DataTypeEnum
  
' The one data type that is ridiculously inefficient in this _
  process is XML. It's returned as binary, and there's no _
  easy way to identify it so we can avoid doubling its size _
  when we send it back to the server.

' To consider: what if we just encoded binary data as unicode _
  characters? We could potentially avoid doing the encoding _
  ourselves, and save on the size of the strings we send back _
  and forth. The only problem would be that we'd need to handle _
  different database collations, etc., which may require us to _
  know about the target database in advance. Something to _
  think about.
  
  Set fi = New FieldInfo
  
  scale_ = Null
  precision_ = Null
  size_ = Null
  
  
  With fld
    functionalTypes = .Type
    If functionalTypes = adBinary Then
      If mGuessBinaryTypes Then
        functionalTypes = GuessBinaryType(fld)
      End If
    End If
    
    Select Case functionalTypes
      Case ADODB.DataTypeEnum.adBigInt
        sqlType = "bigint"
      Case ADODB.DataTypeEnum.adBinary
        sqlType = "binary"
      Case ADODB.DataTypeEnum.adBoolean
        sqlType = "bit"
      Case ADODB.DataTypeEnum.adBSTR
        sqlType = "nvarchar"
      Case ADODB.DataTypeEnum.adChapter
        Err.Raise vbObjectError + erInvalidAdoDataTypeException, "Ado->Sql", "Invalid ADODB.DataTypeEnum value"
      Case ADODB.DataTypeEnum.adChar
        sqlType = "char"
      Case ADODB.DataTypeEnum.adCurrency
        sqlType = "money"
      Case ADODB.DataTypeEnum.adDate
        sqlType = "date"
      Case ADODB.DataTypeEnum.adDBDate
        sqlType = "date"
      Case ADODB.DataTypeEnum.adDBTime
        sqlType = "time"
      Case ADODB.DataTypeEnum.adDBTimeStamp
      ' Interestingly, the "timestamp" data type is actually binary(8) _
        but since it's stored like that natively, we're going to just _
        leave it as binary and make no effort to decode it.
        If .Type = adBinary Then
          sqlType = "datetimeoffset"
        Else
          sqlType = "datetime2"
          scale_ = .NumericScale
        End If
      Case ADODB.DataTypeEnum.adDecimal
        sqlType = "decimal"
      Case ADODB.DataTypeEnum.adDouble
        sqlType = "float"
      Case ADODB.DataTypeEnum.adEmpty
        Err.Raise vbObjectError + erInvalidAdoDataTypeException, "Ado->Sql", "Invalid ADODB.DataTypeEnum value"
      Case ADODB.DataTypeEnum.adError
        Err.Raise vbObjectError + erInvalidAdoDataTypeException, "Ado->Sql", "Invalid ADODB.DataTypeEnum value"
      Case ADODB.DataTypeEnum.adFileTime
        sqlType = "datetime2"
        scale_ = .NumericScale
      Case ADODB.DataTypeEnum.adGUID
        sqlType = "uniqueidentifier"
      Case ADODB.DataTypeEnum.adInteger
        sqlType = "int"
      Case ADODB.DataTypeEnum.adLongVarBinary
        sqlType = "varbinary"
      Case ADODB.DataTypeEnum.adLongVarChar
        sqlType = "varchar"
      Case ADODB.DataTypeEnum.adLongVarWChar
        sqlType = "nvarchar"
      Case ADODB.DataTypeEnum.adNumeric
        sqlType = "numeric"
      Case ADODB.DataTypeEnum.adPropVariant
        sqlType = "sql_variant"
      Case ADODB.DataTypeEnum.adSingle
        sqlType = "real"
      Case ADODB.DataTypeEnum.adSmallInt
        sqlType = "smallint"
      Case ADODB.DataTypeEnum.adTinyInt
        sqlType = "tinyint"
      Case ADODB.DataTypeEnum.adUnsignedBigInt
        sqlType = "binary"
      Case ADODB.DataTypeEnum.adUnsignedInt
        sqlType = "binary"
      Case ADODB.DataTypeEnum.adUnsignedSmallInt
        sqlType = "binary"
      Case ADODB.DataTypeEnum.adUnsignedTinyInt
        sqlType = "binary"
      Case ADODB.DataTypeEnum.adUserDefined
        Err.Raise vbObjectError + erInvalidAdoDataTypeException, "Ado->Sql", "Invalid ADODB.DataTypeEnum value"
      Case ADODB.DataTypeEnum.adVarBinary
        sqlType = "varbinary"
      Case ADODB.DataTypeEnum.adVarChar
        sqlType = "varchar"
      Case ADODB.DataTypeEnum.adVariant
        sqlType = "sql_variant"
      Case ADODB.DataTypeEnum.adVarNumeric
        sqlType = "numeric"
      Case ADODB.DataTypeEnum.adVarWChar
        sqlType = "nvarchar"
      Case ADODB.DataTypeEnum.adWChar
        sqlType = "nchar"
    End Select
    
    Select Case functionalTypes
    ' ADODB.DataTypeEnum.adSingle can be ignored because "real" is a synonym for float(24)
      Case ADODB.DataTypeEnum.adDouble
        precision_ = .Precision
      Case ADODB.DataTypeEnum.adDecimal _
          , ADODB.DataTypeEnum.adNumeric _
          , ADODB.DataTypeEnum.adVarNumeric
        precision_ = .Precision
        scale_ = .NumericScale
      Case ADODB.DataTypeEnum.adBinary _
          , ADODB.DataTypeEnum.adChar _
          , ADODB.DataTypeEnum.adLongVarBinary _
          , ADODB.DataTypeEnum.adLongVarChar _
          , ADODB.DataTypeEnum.adLongVarWChar _
          , ADODB.DataTypeEnum.adVarBinary _
          , ADODB.DataTypeEnum.adVarChar _
          , ADODB.DataTypeEnum.adVarWChar _
          , ADODB.DataTypeEnum.adWChar
        If fld.definedSize > 0 And fld.definedSize <= 8000 Then
          size_ = .definedSize
        Else
          size_ = "max"
        End If
      Case ADODB.DataTypeEnum.adUnsignedBigInt
        size_ = 8
      Case ADODB.DataTypeEnum.adUnsignedInt
        size_ = 4
      Case ADODB.DataTypeEnum.adUnsignedSmallInt
        sqlType = "binary"
        size_ = 2
      Case ADODB.DataTypeEnum.adUnsignedTinyInt
        sqlType = "binary"
        size_ = 1
      Case Else
    End Select
    If scale_ = 255 Then scale_ = 0
    If precision_ = 255 Then precision_ = 0
    
    Nullable = CBool((.Attributes And (ADODB.FieldAttributeEnum.adFldIsNullable Or ADODB.FieldAttributeEnum.adFldMayBeNull)))
  
  If Not IsNull(size_) Then
    fullType = cFormat("%s(%s)", sqlType, size_)
  ElseIf Not IsNull(precision_) And Not IsNull(scale_) Then
    fullType = cFormat("%s(%d,%d)", sqlType, precision_, scale_)
  ElseIf Not IsNull(precision_) Then
    fullType = cFormat("%s(%d)", sqlType, precision_)
  ElseIf IsNull(precision_) And Not IsNull(scale_) Then
    fullType = cFormat("%s(%d)", sqlType, scale_)
  Else
    fullType = sqlType
  End If
  
  Call fi.Setup(.name, fullType, Nullable, fld, functionalTypes)
  
  Set IFieldConverter_GetFieldInfo = fi
    
  End With
End Function



Private Function GuessBinaryType(fld As ADODB.field) As ADODB.DataTypeEnum
  Dim value As Variant
  Dim bytes() As Byte
  GuessBinaryType = adBinary
  If Not mGuessBinaryTypes Then
    Exit Function
  End If
  If fld Is Nothing Then
    Exit Function
  End If
  
  With fld
    If .Type <> adBinary Then
      Exit Function
    End If
  
    On Error GoTo NullError
    value = .value
    If IsNull(value) Then
      Exit Function
    End If
    bytes = value
    
    If IsDbTime(.definedSize, .actualSize, bytes) Then
      GuessBinaryType = adDBTime
      Exit Function
    End If
    If IsDbDateTimeOffset(.definedSize, .actualSize, bytes) Then
      GuessBinaryType = adDBTimeStamp
      Exit Function
    End If
    
  End With
  Exit Function
  
NullError:
End Function

Private Function IsDbTime(definedSize As Long, actualSize As Long, bytes() As Byte) As Boolean
  If definedSize <> 16 Then Exit Function
  If actualSize < 12 Then Exit Function
  Dim comp As LongLong
  Const zeroMask As LongLong = &HFFFFFF00FF00FF00^
  CopyMem VarPtr(comp), VarPtr(bytes(0)), 8
  If (comp And zeroMask) = 0 Then
    If (bytes(0) < 24 And bytes(2) < 60 And bytes(4) < 60) Then
      IsDbTime = True
    End If
  End If
End Function

Private Function IsDbDateTimeOffset(definedSize As Long, actualSize As Long, bytes() As Byte) As Boolean
  If definedSize <> 34 Then Exit Function
  If actualSize < 20 Then Exit Function
  Dim t As DBTIMESTAMPOFFSET
  Dim comp(0 To 2) As LongLong
  Const mask1 As LongLong = &HFFE0FFE0FFF00000^
  Const mask2 As LongLong = &HFFC0FFC0^
  CopyMem VarPtr(comp(0)), VarPtr(bytes(0)), 16
  If (comp(0) And mask1) <> 0 Then
    Exit Function
  End If
  If (comp(1) And mask2) <> 0 Then
    Exit Function
  End If
  CopyMem VarPtr(t), VarPtr(bytes(0)), LenB(t)
  If ((t.month And &HF) <> 0) And ((t.day And &H1F) <> 0) And (t.hour < 24) And (t.minute < 60) And (t.second < 60) And (t.timezone_minute >= -59 And t.timezone_minute <= 59) Then
    IsDbDateTimeOffset = True
    Exit Function
  End If
End Function
Private Property Get IFieldConverter_GuessBinaryTypes() As Boolean
  IFieldConverter_GuessBinaryTypes = mGuessBinaryTypes
End Property

Private Property Let IFieldConverter_GuessBinaryTypes(RHS As Boolean)
  mGuessBinaryTypes = RHS
End Property

Private Property Get IFieldConverter_LeftQuote() As String
  IFieldConverter_LeftQuote = "["
End Property

Private Function IFieldConverter_QuoteName(name As String) As String
  IFieldConverter_QuoteName = IFieldConverter_LeftQuote & Replace(name, IFieldConverter_RightQuote, VBA.String$(2, IFieldConverter_RightQuote)) & IFieldConverter_RightQuote
End Function

Private Property Get IFieldConverter_RightQuote() As String
  IFieldConverter_RightQuote = "]"
End Property

Private Sub IFieldConverterSetup_Setup(formatter As IFormatRules)
  Set mFormatter = formatter
End Sub
