VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlTableConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITableConverter

Private converter As IFieldConverter
Private mFormatter As IFormatRules

Private Sub Class_Initialize()
  Dim converterSetup As IFieldConverterSetup
  Set mFormatter = New SqlFormatRules
  Set converter = New SqlFieldConverter
  Set converterSetup = converter
  converterSetup.Setup mFormatter
End Sub

Private Property Get ITableConverter_Formatter() As IFormatRules
  Set ITableConverter_Formatter = mFormatter
End Property

Public Function GetFields(rs As ADODB.Recordset) As FieldInfo
  Dim fld As ADODB.field
  Dim result As String
  For Each fld In rs.Fields
    If Len(result) > 0 Then result = result & vbCrLf & "  , "
    result = result & converter.GetDeclaration(fld)
  Next

End Function

Private Function ITableConverter_GetFields(rs As ADODB.Recordset) As FieldInfoCollection
  Dim f As ADODB.field
  Dim fi As IFieldInfo
  Set ITableConverter_GetFields = New FieldInfoCollection
  For Each f In rs.Fields
    Set fi = converter.GetFieldInfo(f)
    ITableConverter_GetFields.Add fi
  Next
End Function
