VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Public Property Get FieldConverter() As IFieldConverter
'
'End Property

'
'Public Function GetCreateTable(TableName As String, rs As ADODB.Recordset) As String
'
'End Function

Public Property Get formatter() As IFormatRules

End Property

Public Function GetFields(rs As ADODB.Recordset) As FieldInfoCollection
  
End Function

Public Function GetFactory() As IFactory
End Function
