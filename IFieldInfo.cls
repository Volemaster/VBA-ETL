VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFieldInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Enum FieldInfoOptionsEnum
  fioDeclaration
  fioNameOnly
End Enum

Public Property Get FieldName() As String

End Property

Public Property Get NativeType() As String

End Property

Public Property Get AdoField() As ADODB.field

End Property

Public Property Get Nullable() As Boolean

End Property

Public Function ToString(formatter As IFormatRules, Optional formatOption As FieldInfoOptionsEnum = FieldInfoOptionsEnum.fioNameOnly) As String
  
End Function

Public Function FormatValue(formatter As IFormatRules) As String

End Function

Public Function AppendValue(formatter As IFormatRules, target As IStringBuilder)

End Function

Public Property Get BinaryType() As ADODB.DataTypeEnum
End Property

Public Property Let BinaryType(value As ADODB.DataTypeEnum)
End Property
