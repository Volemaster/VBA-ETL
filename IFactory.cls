VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function NewTempTable(Optional cn As ADODB.Connection, Optional tableConverter As ITableConverter, Optional automaticTearDown As Boolean = False) As ITempTable
End Function

Public Function NewTableWriter() As ITableWriter
End Function
