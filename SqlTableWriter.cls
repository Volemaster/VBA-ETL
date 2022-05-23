VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlTableWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITableWriter

Private mCn As ADODB.Connection
Private mCmd As ADODB.Command
Private mQuotedTableName As String
Private mTableConverter As ITableConverter
Private mFieldProfiles As FieldInfoCollection
Private mPrintStatus As Boolean

Public Event SentRecords(recordCount As Long)
Public Event Waiting()
Public Event LoadComplete(recordCount As Long)

Public Property Get AsITableWriter() As ITableWriter
  Set AsITableWriter = Me
End Property

Public Property Let PrintStatus(value As Boolean)
  mPrintStatus = value
End Property

Public Property Get PrintStatus() As Boolean
  PrintStatus = mPrintStatus
End Property

Public Sub Initialize(cn As ADODB.Connection, tableConverter As ITableConverter, quotedTableName As String, Optional fieldProfiles As FieldInfoCollection)
  Set mCn = cn
  Set mCmd = New ADODB.Command
  Set mCmd.ActiveConnection = mCn
  Set mTableConverter = tableConverter
  Set mFieldProfiles = fieldProfiles
  mQuotedTableName = quotedTableName
End Sub


Private Sub ITableWriter_Delete()
  mCmd.CommandText = "DELETE FROM " & mQuotedTableName
End Sub

Private Sub ITableWriter_Insert(rs As ADODB.Recordset)
  Dim formatter As IFormatRules
  Dim i As Long, maxFieldNum As Long
  Dim stringBuilder As IStringBuilder
  Dim c As FieldInfoCollection
  Dim itemCounter As Long
  Dim insertString As String
  Dim recordsInBatch As Long
  
  Set formatter = mTableConverter.formatter
  If mFieldProfiles Is Nothing Then
    Set c = mTableConverter.GetFields(rs)
  Else
    Set c = mFieldProfiles
  End If
  
  maxFieldNum = c.Count - 1
  Set stringBuilder = New NetStringBuilder
  
  insertString = ";" & vbCrLf & "INSERT INTO " & mQuotedTableName & " (" & _
    c.ToString(formatter, fioNameOnly, ",") & ")" & vbCrLf & _
    "VALUES ("
  
  If rs Is Nothing Then
    'Error
    Exit Sub
  End If
  mCmd.CommandTimeout = 300
  
  With rs
    If .State = ADODB.ObjectStateEnum.adStateClosed Then
      'Error
      Exit Sub
    End If
    itemCounter = 0
    Do Until .EOF
      If itemCounter Mod 1000 = 0 Then
        With mCmd
          If (.State And ADODB.ObjectStateEnum.adStateExecuting) = 0 Then
            If stringBuilder.Length > 0 Then
              If mPrintStatus Then Debug.Print "Send : " & Now()
              .CommandText = stringBuilder.ToString()
              RaiseEvent SentRecords(recordsInBatch)
              .Execute , , ADODB.ExecuteOptionEnum.adExecuteNoRecords Or ADODB.ExecuteOptionEnum.adAsyncExecute
              stringBuilder.Clear
              recordsInBatch = 0
            End If
          End If
        End With
        If mPrintStatus Then Debug.Print Right(Space(4) & itemCounter, 5) & ": " & Now()
        DoEvents
        stringBuilder.Add insertString
      Else
        stringBuilder.Add vbCrLf & ",("
      End If
      For i = 0 To maxFieldNum
        Call c.Item(i).AppendValue(formatter, stringBuilder)
        If i < maxFieldNum Then
          stringBuilder.Add ","
        End If
      Next
      stringBuilder.Add ")"
      itemCounter = itemCounter + 1
      recordsInBatch = recordsInBatch + 1
      .MoveNext
    Loop
    With mCmd
      If .State = ADODB.ObjectStateEnum.adStateExecuting Then
        RaiseEvent Waiting
        If mPrintStatus Then Debug.Print "Wait : " & Now()
      End If
      Do Until (.State And ADODB.ObjectStateEnum.adStateExecuting) = 0
        DoEvents
      Loop
      If stringBuilder.Length > 0 Then
        RaiseEvent SentRecords(recordsInBatch)
        If mPrintStatus Then Debug.Print "Send : " & Now()
        .CommandText = stringBuilder.ToString()
        .Execute , , ADODB.ExecuteOptionEnum.adExecuteNoRecords Or ADODB.ExecuteOptionEnum.adAsyncExecute
      End If
      Do Until (.State And ADODB.ObjectStateEnum.adStateExecuting) = 0
        RaiseEvent Waiting
        DoEvents
      Loop
      RaiseEvent LoadComplete(itemCounter)
      If mPrintStatus Then Debug.Print "Done : " & Now()
    End With
  End With
End Sub

Private Sub Submit()
  
End Sub
