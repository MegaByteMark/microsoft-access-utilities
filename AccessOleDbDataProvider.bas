' Module: SQLDataAccess.bas
'WIP
Option Explicit

' Requires Microsoft ActiveX Data Objects Library (e.g., ADO 6.1) in Tools -> References

' *** IMPORTANT:  This .bas approach is NOT recommended for production.  It's shown for demonstration of basic concepts, but it lacks proper error handling, scalability, and maintainability.  A proper DLL (using .NET or C++) is strongly preferred. ***

' --- Connection String ---
' Store this securely (e.g., in a separate configuration file or registry).
' DO NOT hardcode connection strings directly in your VBA code for production.
Private Const CONN_STRING As String = "Provider=SQLOLEDB.1;Data Source=YourServerName;Initial Catalog=YourDatabaseName;User ID=YourUsername;Password=YourPassword;" ' *** REPLACE with your actual connection details ***

' --- Execute a SQL query with parameters (SQL Injection Safe) ---
Private Function ExecuteSQLWithParams(sql As String, params As Variant) As ADODB.Recordset

    On Error GoTo ErrHandler ' Error handling

    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim i As Long

    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = CONN_STRING
        .CommandText = sql
        .CommandType = adCmdText

        ' Add parameters (important for SQL injection protection)
        If IsArray(params) Then
            For i = LBound(params) To UBound(params)
                .Parameters.Append .CreateParameter("@Param" & i, adVarChar, adParamInput, 255, params(i)) ' Adjust data type and size as needed
            Next i
        End If

        Set rs = .Execute ' Execute the command
    End With

    Set ExecuteSQLWithParams = rs
    Set cmd = Nothing ' Clean up
    Exit Function ' Successful execution

ErrHandler:
    ' Handle errors (log, display message, etc.)
    MsgBox "Database Error: " & Err.Description & " (Error Number: " & Err.Number & ") in ExecuteSQLWithParams", vbCritical, "Database Error"
    ' Optionally, you could return an empty recordset or Null here to indicate an error.
    Set ExecuteSQLWithParams = Nothing ' Or Set ExecuteSQLWithParams = New ADODB.Recordset if you want an empty recordset
    Set cmd = Nothing
End Function

' --- Get a single value from a query ---
Public Function GetValue(sql As String, params As Variant) As Variant

    On Error GoTo ErrHandler

    Dim rs As ADODB.Recordset

    Set rs = ExecuteSQLWithParams(sql, params)

    If Not rs Is Nothing Then ' Check if the recordset is valid (no error)
        If Not rs.EOF Then
            GetValue = rs.Fields(0).Value ' Get the first field's value
        Else
            GetValue = Null ' Or some other default value
        End If
    End If 'No Else needed, as if rs is nothing, GetValue will remain at its default value

    If Not rs Is Nothing Then ' Check if the recordset is valid (no error) before closing
        rs.Close
        Set rs = Nothing
    End If

    Exit Function

ErrHandler:
    MsgBox "Database Error: " & Err.Description & " (Error Number: " & Err.Number & ") in GetValue", vbCritical, "Database Error"
    GetValue = Null ' Or some other default value
End Function

' --- Get a DataRow (simulated - returns an array) ---
Public Function GetDataRow(sql As String, params As Variant) As Variant

    On Error GoTo ErrHandler

    Dim rs As ADODB.Recordset
    Dim rowData As Variant

    Set rs = ExecuteSQLWithParams(sql, params)

    If Not rs Is Nothing Then ' Check if the recordset is valid (no error)
        If Not rs.EOF Then
            rowData = rs.GetRows(1) ' Get the first row
            GetDataRow = rowData(0) ' Return the array of values (simulating a row)
        Else
            GetDataRow = Null
        End If
    End If

    If Not rs Is Nothing Then ' Check if the recordset is valid (no error) before closing
        rs.Close
        Set rs = Nothing
    End If

    Exit Function

ErrHandler:
    MsgBox "Database Error: " & Err.Description & " (Error Number: " & Err.Number & ") in GetDataRow", vbCritical, "Database Error"
    GetDataRow = Null
End Function

' --- Get a Recordset ---
Public Function GetRecordset(sql As String, params As Variant) As ADODB.Recordset
    On Error GoTo ErrHandler
    Set GetRecordset = ExecuteSQLWithParams(sql, params) ' Directly return the recordset
    Exit Function

ErrHandler:
    MsgBox "Database Error: " & Err.Description & " (Error Number: " & Err.Number & ") in GetRecordset", vbCritical, "Database Error"
    Set GetRecordset = Nothing
End Function

' --- Example Usage (in another module) ---
Sub TestSQLFunctions()

    Dim sql As String
    Dim params As Variant
    Dim rs As ADODB.Recordset
    Dim value As Variant
    Dim row As Variant

    ' --- Get a single value ---
    sql = "SELECT COUNT(*) FROM YourTable WHERE SomeColumn = ?"
    params = Array("SomeValue") ' Parameterized query
    value = GetValue(sql, params)
    Debug.Print "Count:", value

    ' --- Get a row ---
    sql = "SELECT * FROM YourTable WHERE ID = ?"
    params = Array(123)
    row = GetDataRow(sql, params)
    If Not IsNull(row) Then
        Debug.Print "Row Data:", Join(row, ", ") ' Print the array of row values
    End If

    ' --- Get a recordset ---
    sql = "SELECT * FROM YourTable WHERE SomeOtherColumn = ?"
    params = Array("AnotherValue")
    Set rs = GetRecordset(sql, params)
    If Not rs Is Nothing Then 'Check if rs is valid before trying to use it
        If Not rs.EOF Then
            ' Process the recordset
            Do While Not rs.EOF
                Debug.Print rs!Column1, rs!Column2 ' ... etc.
                rs.MoveNext
            Loop
        End If
        rs.Close
        Set rs = Nothing
    End If

End Sub
