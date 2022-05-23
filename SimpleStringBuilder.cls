VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mValue As String
Private mSize As Long
Private mLength As Long
Private mChunkSize As Long

Implements IStringBuilder

Private Sub IStringBuilder_Clear()
  mLength = 0
End Sub

Private Property Get IStringBuilder_Length() As Long
  IStringBuilder_Length = mLength
End Property

Private Sub Class_Initialize()
  mChunkSize = &H4000000
  mLength = 0
End Sub

Private Function IStringBuilder_Add(value As String) As IStringBuilder
  Dim l As Long
  l = Len(value)
  If mLength + l > mSize Then
    mSize = mSize + mChunkSize
    mValue = mValue & Space$(mSize - mLength)
  End If
  Mid$(mValue, mLength + 1, l) = value
  mLength = mLength + l
  Set IStringBuilder_Add = Me
End Function

Private Function IStringBuilder_ToString() As String
  IStringBuilder_ToString = Left$(mValue, mLength)
End Function

