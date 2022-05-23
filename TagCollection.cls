VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TagCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTags() As String
Private mTagCount As Long
Private mMaxSize As Long
Private Const SIZESTEP As Long = 5

Private Sub Class_Initialize()
'  ReDim mTags(0 To SIZESTEP - 1)
  mMaxSize = 0
  mTagCount = 0
End Sub

Public Function HasTag(ParamArray Tags() As Variant) As Boolean
  If mTagCount = 0 Then Exit Function
  Dim v As Variant, i As Long
  Dim pa As Variant
  If UBound(Tags) >= LBound(Tags) Then
    If VarType(Tags(LBound(Tags))) And vbArray <> 0 Then
      pa = Tags(0)
    End If
  Else
    pa = Tags
  End If
  For Each v In pa
    For i = 0 To mTagCount - 1
      If mTags(i) = v Then
        HasTag = True
        Exit Function
      End If
    Next
  Next
End Function

Public Function Copy() As TagCollection
  Set Copy = New TagCollection
  If mTagCount = 0 Then Exit Function
  Copy.MaxSize = MaxSize
  Dim i As Long
  For i = 0 To mTagCount - 1
    Copy.AddTag mTags(i)
  Next
End Function

Public Sub RemoveTag(tag As String)
  Dim i As Long
  Dim adjust As Long
  For i = 0 To mTagCount - 1
    If adjust <> 0 Then
      mTags(i + adjust) = mTags(i)
    End If
    If mTags(i) = tag Then
      adjust = -1
    End If
  Next
  mTagCount = mTagCount + adjust
End Sub

Public Property Get MaxSize() As Long
  MaxSize = mMaxSize
End Property

Public Property Let MaxSize(value As Long)
  mMaxSize = value
  If mTagCount > mMaxSize Then
    mTagCount = mMaxSize
  End If
  ReDim Preserve mTags(0 To mMaxSize - 1)
End Property

Public Sub AddTag(tag As String)
  Dim i As Long
  For i = 0 To mTagCount - 1
    If mTags(i) = tag Then Exit Sub
  Next
  mTagCount = mTagCount + 1
  If mTagCount >= mMaxSize Then
    MaxSize = mMaxSize + SIZESTEP
  End If
  mTags(mTagCount - 1) = tag
End Sub
