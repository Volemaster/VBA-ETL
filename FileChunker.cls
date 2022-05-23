VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileChunker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSections As Scripting.Dictionary
Private mLines As Scripting.Dictionary

Private Sub Class_Initialize()
  Set mSections = New Scripting.Dictionary
End Sub

Public Function GetSection(sectionName As String) As String
  Dim sectionChunk As VBA.Collection
  Dim v As Variant
  If mSections.Exists(sectionName) Then
    Set sectionChunk = mSections(sectionName)
    For Each v In sectionChunk
      GetSection = GetSection & v
    Next
  End If
End Function

Public Function GetWithoutSections(ParamArray sectionsToExclude() As Variant)
  Dim i As Long, tt As TaggedText
  Dim sb As IStringBuilder
  On Error GoTo HandleError
  Set sb = New NetStringBuilder
  For i = 1 To mLines.Count
    Set tt = mLines(i)
    If Not tt.Tags.HasTag(sectionsToExclude) Then
      sb.Add tt.Text
      sb.Add vbCrLf
    End If
  Next
  GetWithoutSections = sb.ToString()
  Exit Function
HandleError:
  Stop
  Resume
End Function



Public Sub Load(filePath As String)
  ParseFile filePath
End Sub

Private Sub ParseFile(filePath As String)
  Dim fso As Scripting.FileSystemObject, f As Scripting.File, ts As Scripting.TextStream
  Dim chunks As VBA.Collection, txt As IStringBuilder, thisLine As String
  Dim regex As VBScript_RegExp_55.RegExp, m As VBScript_RegExp_55.Match
  Dim sectionName As String, sectionType As String
  Dim tagged As TaggedText
  Dim lineNumber As Long
  
  Set fso = New Scripting.FileSystemObject
  If Not fso.FileExists(filePath) Then
    Err.Raise 53, , "The """ & filePath & """ was not found or is not accessible."
    Exit Sub
  End If
  Set f = fso.GetFile(filePath)
  Set ts = f.OpenAsTextStream(ForReading)
  Set txt = New NetStringBuilder
  Set regex = New VBScript_RegExp_55.RegExp
  With regex
    .MultiLine = False
    .IgnoreCase = True
    ' Begin/End | Section Name
    .Pattern = "\s*--!\s+(Begin|End)\s+Section:\s+(.+)"
  End With
  
  Set mLines = New Scripting.Dictionary
  Set tagged = New TaggedText
  With ts
    Do Until .AtEndOfStream
      thisLine = .ReadLine()
      lineNumber = lineNumber + 1
      tagged.Text = thisLine
      If regex.test(thisLine) Then
        Set m = regex.Execute(thisLine).Item(0)
        sectionType = m.SubMatches(0)
        sectionName = m.SubMatches(1)
        If sectionType = "Begin" Then
          tagged.Tags.AddTag sectionName
        End If
      End If
      mLines.Add lineNumber, tagged
      Set tagged = tagged.NextLine()
      If sectionType = "End" Then
        tagged.Tags.RemoveTag sectionName
      End If
    Loop
    .Close
  End With
End Sub
