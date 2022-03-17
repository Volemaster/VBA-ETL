VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlTempTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITempTable
Implements ITempTableSetup

Private mCn As ADODB.Connection
Private mTableName As Variant
Private mAutomaticTearDown As Boolean
Private mConverter As ITableConverter
Private mFieldProfiles   As FieldInfoCollection

Private Property Get TempTable() As ITempTable
  Set TempTable = Me
End Property

Private Sub Class_Terminate()
  If mAutomaticTearDown Then
    TempTable.Drop
  End If
End Sub

Private Property Let ITempTable_AutomaticTearDown(RHS As Boolean)
  mAutomaticTearDown = RHS
End Property

Private Property Get ITempTable_AutomaticTearDown() As Boolean
  ITempTable_AutomaticTearDown = mAutomaticTearDown
End Property

Private Property Get ITempTable_Connection() As ADODB.Connection
  Set ITempTable_Connection = mCn
End Property

Private Property Get ITempTable_Converter() As ITableConverter
  Set ITempTable_Converter = mConverter
End Property

Private Sub ITempTable_Create(tableName As String, fieldProfiles As FieldInfoCollection)
  Dim v As Variant
  mTableName = tableName
  Set mFieldProfiles = fieldProfiles
  v = GetCreateSql()
  If Not (IsNull(v)) Then
    mCn.Execute CStr(v)
  End If
End Sub

Private Sub ITempTable_Drop()
  If IsNull(mTableName) Then Exit Sub
  If mCn Is Nothing Then Exit Sub
  If mCn.State = ADODB.ObjectStateEnum.adStateClosed Then Exit Sub

  Dim cmd As ADODB.Command, sql As Variant
  Set cmd = New ADODB.Command
  With cmd
    Set .ActiveConnection = mCn
    sql = GetDropText()
    If IsNull(sql) Then Exit Sub
    .CommandText = sql
    .Execute
  End With
End Sub

Public Function GetCreateSqlExplicit(fieldList As String) As Variant
  If IsNull(mTableName) Then
    GetCreateSqlExplicit = Null
  Else
    GetCreateSqlExplicit = "CREATE TABLE " & mConverter.formatter.QuoteName(CStr(mTableName)) & vbCrLf & "(" & vbCrLf & _
      fieldList & ");" & vbCrLf
  End If
End Function

Public Sub Populate(rs As ADODB.Recordset)
If rs Is Nothing Then
  Err.Raise 424, "SqlTempTable.Populate", "Cannot populate temporary table from an uninitialized recordset."
  Exit Sub
End If
If rs.State = ADODB.ObjectStateEnum.adStateClosed Then
  Err.Raise 3704, "SqlTempTable.Populate", "Cannot populate temporary table from a closed recordset."
  Exit Sub
End If
Dim fic As FieldInfoCollection
Set fic = mConverter.GetFields(rs)



End Sub

Public Function GetCreateSql() As Variant
  Dim fieldList As String
  Dim formatter As IFormatRules
  Dim i As Long
  Dim sql As String
  If IsNull(mTableName) Then
    GetCreateSql = Null
    Exit Function
  End If
  If IsNull(mFieldProfiles) Then
    Err.Raise vbObjectError + erNotProperlyInitializedException, "SqlTempTable.GetCreateSql", "You must initialize this object by providing a FieldInfoCollection to ITempTable_Create() before calling this function."
    Exit Function
  End If
  With mConverter
    Set formatter = .formatter
  End With
  With mFieldProfiles
    If .Count = 0 Then
      GetCreateSql = Null
      Exit Function
    End If
    sql = .ToString(formatter, fioDeclaration, vbCrLf & ",")
    GetCreateSql = CStr(GetDropText()) & vbCrLf & _
      "CREATE TABLE " & formatter.QuoteName(CStr(mTableName)) & " (" & vbCrLf & sql & ");"
  End With
End Function

Public Function GetDropText() As Variant
  If IsNull(mTableName) Then
    GetDropText = Null
  Else
    GetDropText = "DROP TABLE IF EXISTS " & mConverter.formatter.QuoteName(CStr(mTableName)) & ";" & vbCrLf
  End If
End Function

Private Property Get ITempTable_TableName() As Variant
  ITempTable_TableName = mTableName
End Property

Private Sub ITempTableSetup_Setup(cn As ADODB.Connection, tableConverter As ITableConverter, Optional withAutomaticTearDown As Boolean = True)
  Set mCn = cn
  Set mConverter = tableConverter
  mAutomaticTearDown = withAutomaticTearDown
End Sub
