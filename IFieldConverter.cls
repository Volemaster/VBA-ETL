VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFieldConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetDeclaration(FieldInfo As IFieldInfo) As String 'fld As ADODB.Field) As String

End Function

Public Function GetFieldInfo(fld As ADODB.field) As IFieldInfo

End Function

Public Property Get formatter() As IFormatRules

End Property

Public Property Get GuessBinaryTypes() As Boolean

End Property

Public Property Let GuessBinaryTypes(value As Boolean)

End Property
