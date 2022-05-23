VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaggedText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mText As String
Dim mTagCollection As TagCollection

Public Property Get Text() As String
  Text = mText
End Property

Public Property Let Text(value As String)
  mText = value
End Property

Public Property Get Tags() As TagCollection
  If mTagCollection Is Nothing Then Set mTagCollection = New TagCollection
  Set Tags = mTagCollection
End Property

Public Property Set Tags(value As TagCollection)
  Set mTagCollection = value
End Property

Public Function NextLine() As TaggedText
  Set NextLine = New TaggedText
  Set NextLine.Tags = Tags.Copy()

End Function
