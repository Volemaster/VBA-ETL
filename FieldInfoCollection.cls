VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldInfoCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mStepSize As Long = 8
Private mFields() As IFieldInfo
Private mValidIndex As Long
Private mSize As Long

Private Sub Class_Initialize()
  mSize = mStepSize * 4
  mValidIndex = -1
  ReDim mFields(0 To mSize - 1)
End Sub

Public Sub Add(field As IFieldInfo)
  mValidIndex = mValidIndex + 1
  If mValidIndex >= mSize Then
    mSize = mSize + mStepSize
    ReDim Preserve mFields(0 To mSize - 1)
  End If
  Set mFields(mValidIndex) = field
End Sub

Public Property Get Items() As IFieldInfo()
  Dim theseItems() As IFieldInfo
  ReDim theseItems(0 To mValidIndex) As IFieldInfo
  Dim i As Long
  If Me.Count = 0 Then
    Exit Property
  End If
  For i = 0 To mValidIndex
    Set theseItems(i) = mFields(i)
  Next
  Items = theseItems
End Property

Public Property Get Item(Index As Long) As IFieldInfo
  If Index > mValidIndex Then
    Err.Raise 440, "FieldCollection", "The index vale (" & Index & ") supplied to FieldCollection is outside of bounds."
  ElseIf Index >= 0 Then
    Set Item = mFields(Index)
  Else
    Err.Raise 440, "FieldCollection", "The index vale (" & Index & ") supplied to FieldCollection is outside of bounds."
  End If
End Property

Public Property Get Count() As Long
  Count = mValidIndex + 1
End Property

Public Function ToString(formatRules As IFormatRules, Optional formatOption As FieldInfoOptionsEnum = FieldInfoOptionsEnum.fioDeclaration, Optional fieldSeparator As String = vbCrLf) As String
  Dim i As Long
  Dim fi As IFieldInfo
  Dim result As String
  If Count = 0 Then Exit Function
  For i = 0 To Count - 1
    Set fi = mFields(i)
    result = result & fi.ToString(formatRules, formatOption) & fieldSeparator
  Next
  ToString = Left$(result, Len(result) - Len(fieldSeparator))
End Function
