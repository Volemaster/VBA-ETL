VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NetStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mStringBuilder As Object
' Or add a reference to mscorlib and use: _
Private mStringBuilder As mscorlib.stringBuilder

Implements IStringBuilder

Private Sub Class_Initialize()
  Set mStringBuilder = CreateObject("System.Text.StringBuilder")
  ' Or use this if you added mscorlib as a reference: _
  Set mStringBuilder = New mscorlib.stringBuilder
  mStringBuilder.Capacity = &H4000000
End Sub

Private Function IStringBuilder_Add(value As String) As IStringBuilder
  mStringBuilder.Append_3 value
  Set IStringBuilder_Add = Me
End Function

Private Sub IStringBuilder_Clear()
  With mStringBuilder
    .Clear
    .Capacity = &H4000000
  End With
End Sub

Private Property Get IStringBuilder_Length() As Long
  IStringBuilder_Length = mStringBuilder.Length
End Property

Private Function IStringBuilder_ToString() As String
  IStringBuilder_ToString = mStringBuilder.ToString()
  mStringBuilder.Append_3 IStringBuilder_ToString
End Function

