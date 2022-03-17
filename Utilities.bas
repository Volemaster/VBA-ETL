Attribute VB_Name = "Utilities"
Option Explicit

' All of the scalar, system-defined SQL Server data types:
Public Const SqlServerTypeTestSql As String = _
"SELECT CAST(NULL AS image) AS [image], CAST(NULL AS text) AS [text], CAST(NEWID() AS uniqueidentifier) AS [uniqueidentifier], CAST(GETDATE() AS date) AS [date], CAST(GETDATE() AS time) AS [time], CAST(GETDATE() AS datetime2) AS [datetime2], CAST(GETDATE() AS datetimeoffset) AS [datetimeoffset], CAST(1.00 AS tinyint) AS [tinyint], CAST(1.00 AS smallint) AS [smallint], CAST(1.00 AS int) AS [int], CAST(GETDATE() AS smalldatetime) AS [smalldatetime], CAST(1.00 AS real) AS [real], CAST(1.00 AS money) AS [money], CAST(GETDATE() AS datetime) AS [datetime], CAST(1.00 AS float) AS [float], CAST(NULL AS sql_variant) AS [sql_variant], CAST(NULL AS ntext) AS [ntext], CAST(1.00 AS bit) AS [bit], CAST(1.00 AS decimal(17,4)) AS [decimal], CAST(1.00 AS numeric) AS [numeric], CAST(1.00 AS smallmoney) AS [smallmoney], CAST(1.00 AS bigint) AS [bigint], CAST(NULL AS hierarchyid) AS [hierarchyid], CAST(NULL AS geometry) AS [geometry], CAST(NULL AS geography) AS [geography], CAST(NULL AS varbinary(200)) AS [varbinary], " & _
  "CAST(NULL AS varchar(200)) AS [varchar], CAST(NULL AS binary(200)) AS [binary], CAST(NULL AS char(200)) AS [char], CAST(GETDATE() AS timestamp) AS [timestamp], CAST(NULL AS nvarchar(200)) AS [nvarchar], CAST(NULL AS nchar(200)) AS [nchar], CAST('<xml></xml>' AS xml) AS [xml], CAST(NULL AS sysname) AS [sysname], CAST(CAST(GETDATE() AS Time) AS Varbinary(12)) AS Time2 "
Public Const SqlServerTypeTestSql2 As String = _
    "SET NOCOUNT ON;" & vbCrLf & _
    "DECLARE @SQL NVarchar(Max)" & vbCrLf & _
    "SET @SQL =" & vbCrLf & _
    "'SELECT " & vbCrLf & _
    "   ' + STUFF(" & vbCrLf & _
    "(SELECT" & vbCrLf & _
    "  CONCAT(CHAR(13), '  , CAST(', CASE WHEN precision >0 OR scale >0 THEN '1.00' ELSE 'NULL' END, ' AS ', name, CASE WHEN precision>0 or scale>0 THEN '' WHEN max_length BETWEEN 4000 AND 8000 THEN '(200)' ELSE '' END,') AS ', QUOTENAME(name))" & vbCrLf & _
    " FROM sys.types where is_table_type=0 and is_user_defined=0" & vbCrLf & _
    "FOR XML PATH(''),Type).value('.','nvarchar(max)'),1,4,'') + ' WHERE 1=0'" & vbCrLf & _
    "EXEC sp_executesql @SQL;"

Public Declare PtrSafe Sub CopyMem Lib "ntdll.dll" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal size As Long)

Public Declare PtrSafe Sub CopyMemory Lib "ntdll.dll" Alias _
    "RtlMoveMemory" (Destination As Any, Source As Any, _
    ByVal Length As Long)


Public Function GetDisconnectedRecordset(rs As ADODB.Recordset) As ADODB.Recordset
  Dim stream As ADODB.stream
  Set stream = New ADODB.stream
  rs.Save stream, adPersistADTG
  Set GetDisconnectedRecordset = New ADODB.Recordset
  With GetDisconnectedRecordset
    .LockType = adLockOptimistic
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open stream
  End With
End Function


Public Function cFormat(txt As String, ParamArray values())
  Dim v As Variant
  Dim nextNumber As Long, nextString As Long
  cFormat = txt
  For Each v In values
    nextNumber = InStr(1, cFormat, "%d", vbTextCompare)
    nextString = InStr(1, cFormat, "%s", vbTextCompare)
    If (nextNumber > 0 Or nextString > 0) Then
      If nextNumber > 0 And (nextNumber < nextString Or nextString = 0) Then
        If VBA.IsNumeric(v) Then
          cFormat = Replace(cFormat, "%d", v, 1, 1, vbTextCompare)
        Else
          Err.Raise erInvalidNumberFormat, "cFormat", "Invalid number format"
        End If
      ElseIf nextString > 0 Then
        cFormat = Replace(cFormat, "%s", v, 1, 1, vbTextCompare)
      End If
    Else
      Exit For
    End If
  Next
End Function


