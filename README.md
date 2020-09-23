<div align="center">

## ACCESS / ORACLE ADO Connection and Recordset


</div>

### Description

To Provide a Recordset Template

(view,execute,edit,Add new)

To Provide a Database Connection Template (ORACLE

and MS Access)
 
### More Info
 
Add Microsoft ActiveX Data Object 2.X library

into Project Reference.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Herman Ibrahim](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/herman-ibrahim.md)
**Level**          |Intermediate
**User Rating**    |4.6 (65 globes from 14 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/herman-ibrahim-access-oracle-ado-connection-and-recordset__1-6321/archive/master.zip)





### Source Code

```
'Put This part in a module
'======================================
'Starrt of Module
'======================================
Option Explicit
Public Enum RSMethod
 VIEW_RECORD = 0
 EDIT_RECORD = 1
 EXEC_SQL = 2
 NEW_RECORD = 3
End Enum
Function dbConnection(strDatabaseType As String, strDBService As String, Optional strUserID As String, Optional strPassword As String) As ADODB.Connection
 Dim objDB As New ADODB.Connection
 Dim strConnectionString As String
 If strDatabaseType = "ORACLE" Then
 'Define ORACLE database connection string
 strConnectionString = "Driver={Microsoft ODBC Driver for Oracle};ConnectString=" & strDBService & ";UID=" & strUserID & ";PWD=" & strPassword & ";"
 ElseIf strDatabaseType = "MSACCESS" Then
 'Define Microsoft Access database connection string
 strConnectionString = "DBQ=" & strDBService
 strConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)}; " & strConnectionString
 End If
 With objDB
 .Mode = adModeReadWrite ' connection mode ???
 .ConnectionTimeout = 10 'Indicates how long to wait while establishing a connection before terminating the attempt and generating an error.
 .CommandTimeout = 5 ' seconds given to execute any command
 .CursorLocation = adUseClient ' use the appropriate cursor ???
 .Open strConnectionString 'open the database connection
 End With
 Set dbConnection = objDB
End Function
Function CreateRecordSet(ByRef dbConn As ADODB.Connection, ByRef rs As ADODB.Recordset, ByVal method As RSMethod, Optional strSQL As String, Optional TableName As String) As ADODB.Recordset
' close the recordset first if it's open...
' otherwise an error will occured
'(open a recordset which is already opened...)
if rs.State=1 then
rs.close
end if
 Select Case method
 Case RSMethod.NEW_RECORD
 rs.ActiveConnection = dbConn
 rs.CursorType = adOpenKeyset
 rs.LockType = adLockOptimistic
 rs.CursorLocation = adUseServer
 rs.Open TableName
 Case RSMethod.EDIT_RECORD
 rs.ActiveConnection = dbConn
 rs.Source = strSQL
 rs.CursorType = adOpenKeyset
 rs.LockType = adLockOptimistic
 rs.CursorLocation = adUseClient
 rs.Open
' Debug.Print "SQL Statement in EDIT Mode (Createrecordset) : " & strSQL
' Debug.Print "Found " & rs.RecordCount & " records"
 Case RSMethod.VIEW_RECORD
 rs.ActiveConnection = dbConn 'dbConnection 'dbConn
 rs.Source = strSQL
 rs.CursorType = adOpenForwardOnly
 rs.CursorLocation = adUseClient
 rs.Open
' Debug.Print "Found " & rs.RecordCount & " records"
 rs.ActiveConnection = Nothing
 Case RSMethod.EXEC_SQL
 Set rs = dbConn.Execute(strSQL)
 End Select
 Set CreateRecordSet = rs
End Function
'======================================
'End Of Module
'======================================
'=================================================
'======================================
'Sample of subroutines...
'======================================
Sub Add_New_Record()
 Dim objRecSet As New ADODB.Recordset
 Dim objConn As New ADODB.Connection
 Dim strUserID As String
 Dim strPassword As String
 Dim strTableName As String
 Dim strDBType As String
 Dim strDBName As String
 strTableName = "YOURTABLE"
 strPassword = "YourPassword"
 strUserID = "YourUserID"
 If strDBType = "MSACCESS" Then
 ' strDBName is your Database Name
 strDBName = App.Path & "\YourAccessDB.mdb"
 ElseIf strDBType = "ORACLE" Then
 ' strDBName is your Oracle Service Name
 strDBName = "YOUR_ORACLE_SERVICE_NAME"
 strTableName = strUserID & "." & strTableName
 'Table name format ::> USERID.TABLENAME
 Else
 MsgBox "Database is other than ORACLE or Microsoft"
 Exit Sub
 End If
 Set objConn = dbConnection(strDBType, strDBName, "userid", "password")
 'send NEW_RECORD and strTableName as a part of parameters
 Set objRecSet = CreateRecordSet(objConn, objRecSet, NEW_RECORD, , strTableName)
 objConn.BeginTrans
 With objRecSet
 .AddNew
 .Fields("FIELD1").Value = "your value1"
 .Fields("FIELD2").Value = "your value2"
 .Fields("FIELD3").Value = "your value3"
 .Fields("FIELD4").Value = "your value4"
 .Fields("FIELD5").Value = "your value5"
 .Update
 End With
 If objConn.Errors.Count = 0 Then
 objConn.CommitTrans
 Else
 objConn.RollbackTrans
 End If
 objRecSet.Close
 objConn.Close
 Set objRecSet = Nothing
 Set objConn = Nothing
End Sub
Sub View_Record_Only()
 Dim strSQL As String
 Dim strDBName As String
 Dim strDBType As String
 Dim strUserID As String
 Dim strPassword As String
 Dim objRecSet As New ADODB.Recordset
 Dim objConn As New ADODB.Connection
 If strDBType = "MSACCESS" Then
 ' strDBName is your Database Name
 strDBName = App.Path & "\YourAccessDB.mdb"
 ElseIf strDBType = "ORACLE" Then
 ' strDBName is your Oracle Service Name
 strDBName = "YOUR_ORACLE_SERVICE_NAME"
 Else
 MsgBox "Database is other than ORACLE or Microsoft"
 Exit Sub
 End If
 strPassword = "YourPassword"
 strUserID = "YourUserID"
 strSQL = "SELECT * from USER_TABLE"
 Set objConn = dbConnection(strDBType, strDBName, "userid", "password")
 'create a disconnected recordset
 Set objRecSet = CreateRecordSet(objConn, objRecSet, VIEW_RECORD, strSQL)
 objConn.Close
 Set objConn = Nothing
 'manipulate the recordset here.....
 'manipulate the recordset here.....
 'manipulate the recordset here.....
 objRecSet.Close
 Set objRecSet = Nothing
End Sub
Sub Edit_Existing_Record()
 Dim objRecSet As New ADODB.Recordset
 Dim objConn As New ADODB.Connection
 Dim strUserID As String
 Dim strPassword As String
 Dim strSQL As String
 Dim strDBType As String
 Dim strDBName As String
 strTableName = "YOURTABLE"
 strPassword = "YourPassword"
 strUserID = "YourUserID"
 If strDBType = "MSACCESS" Then
 ' strDBName is your Database Name
 strDBName = App.Path & "\YourAccessDB.mdb"
 ElseIf strDBType = "ORACLE" Then
 ' strDBName is your Oracle Service Name
 strDBName = "YOUR_ORACLE_SERVICE_NAME"
 Else
 MsgBox "Database is other than ORACLE or Microsoft"
 Exit Sub
 End If
 strSQL = "Select * from YOUR_TABLE"
 Set objConn = dbConnection(strDBType, strDBName, "userid", "password")
 'send EDIT_RECORD and strSQL as a part of parameters
 Set objRecSet = CreateRecordSet(objConn, objRecSet, EDIT_RECORD, strSQL)
 With objRecSet
 .Fields("FIELD1").Value = "your value1"
 .Update
 End With
 objRecSet.Close
 objConn.Close
 Set objRecSet = Nothing
 Set objConn = Nothing
End Sub
'======================================
'End of Sample of subroutines...
'======================================
```

