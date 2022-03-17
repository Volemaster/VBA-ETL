VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITempTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get tableName() As Variant

End Property

Public Property Get Connection() As ADODB.Connection

End Property

Public Property Get converter() As ITableConverter

End Property

Public Property Get AutomaticTearDown() As Boolean

End Property

Public Property Let AutomaticTearDown(value As Boolean)

End Property

Public Sub Create(tableName As String, fieldProfiles As FieldInfoCollection)

End Sub

Public Sub Drop()

End Sub


