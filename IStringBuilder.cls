VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Property Get Length() As Long
End Property

Public Function Add(value As String) As IStringBuilder
End Function

Public Function ToString() As String
End Function

Public Sub Clear()
End Sub
