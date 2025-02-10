' MIT License
'
' Copyright (c) 2025 MegaByteMark (https://github.com/MegaByteMark)
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Option Explicit

Private connString As String
Private connTimeout As Integer
Private cmdTimeout As Integer

' Sets the connection string for the database connection.
' @param value The connection string.
Public Property Let ConnectionString(value As String)
    connString = value
End Property

' Gets the connection string for the database connection.
' @return The connection string.
Public Property Get ConnectionString() As String
    ConnectionString = connString
End Property

' Sets the connection timeout for the database connection.
' @param value The connection timeout in seconds.
Public Property Let connectionTimeout(value As Integer)
    connTimeout = value
End Property

' Gets the connection timeout for the database connection.
' @return The connection timeout in seconds.
Public Property Get connectionTimeout() As Integer
    connectionTimeout = connTimeout
End Property

' Sets the command timeout for the database commands.
' @param value The command timeout in seconds.
Public Property Let CommandTimeout(value As Integer)
    cmdTimeout = value
End Property

' Gets the command timeout for the database commands.
' @return The command timeout in seconds.
Public Property Get CommandTimeout() As Integer
    CommandTimeout = cmdTimeout
End Property

' Initializes the SqlServerDataProvider with the specified connection string and timeouts.
' @param connString The connection string.
' @param connTimeout The connection timeout in seconds (optional).
' @param cmdTimeout The command timeout in seconds (optional).
Public Sub Initialize(ByVal connString As String, Optional ByVal connTimeout As Integer = 0, Optional ByVal cmdTimeout As Integer = 0)
    ConnectionString = connString
    connectionTimeout = connTimeout
    CommandTimeout = cmdTimeout
End Sub

' Gets a new database connection.
' @param connectionTimeout The connection timeout in seconds (optional).
' @return An ADODB.Connection object.
Public Function GetDbConnection(Optional ByVal connectionTimeout As Integer = 30) As ADODB.Connection
    Dim conn As ADODB.Connection

    If ConnectionString = "" Then
        Err.Raise 1, "GetDbConnection", "Connection string not set"
    End If

    Set conn = New ADODB.Connection
    conn.ConnectionString = ConnectionString
    conn.connectionTimeout = connectionTimeout
    
    conn.Open
    
    Set GetDbConnection = conn
End Function

' Executes a non-query SQL command.
' @param sql The SQL command to execute.
' @param params A dictionary of parameters (optional).
' @return The number of affected rows.
Public Function ExecuteNonQuery(ByVal sql As String, Optional ByVal params As Scripting.Dictionary) As Long
    Dim conn As ADODB.Connection
    Dim affectedRows As Long

    Set conn = GetDbConnection()
    affectedRows = ExecuteNonQueryOnConnection(conn, sql, params)
    conn.Close

    Set conn = Nothing

    ExecuteNonQuery = affectedRows
End Function

' Executes a non-query SQL command on an existing connection.
' @param conn The existing ADODB.Connection object.
' @param sql The SQL command to execute.
' @param params A dictionary of parameters (optional).
' @return The number of affected rows.
Public Function ExecuteNonQueryOnConnection(ByRef conn As ADODB.Connection, ByVal sql As String, Optional ByVal params As Scripting.Dictionary) As Long
    Dim cmd As ADODB.Command
    Dim affectedRows As Long

    Set cmd = GetDbCommand(conn, sql, params)

    cmd.Execute affectedRows

    Set cmd = Nothing

    ExecuteNonQueryOnConnection = affectedRows
End Function

' Executes a scalar SQL command.
' @param sql The SQL command to execute.
' @param params A dictionary of parameters (optional).
' @return The result of the scalar query.
Public Function ExecuteScalar(ByVal sql As String, Optional ByVal params As Scripting.Dictionary) As Variant
    Dim conn As ADODB.Connection
    Dim value As Variant

    Set conn = GetDbConnection()
    value = ExecuteScalarOnConnection(conn, sql, params)
    conn.Close

    Set conn = Nothing

    ExecuteScalar = value
End Function

' Executes a scalar SQL command on an existing connection.
' @param conn The existing ADODB.Connection object.
' @param sql The SQL command to execute.
' @param params A dictionary of parameters (optional).
' @return The result of the scalar query.
Public Function ExecuteScalarOnConnection(ByRef conn As ADODB.Connection, ByVal sql As String, Optional ByVal params As Scripting.Dictionary) As Variant
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim value As Variant

    Set cmd = GetDbCommand(conn, sql, params)

    Set rs = cmd.Execute

    If Not rs.EOF Then
        value = rs.Fields(0).value
    End If

    rs.Close
    Set rs = Nothing
    Set cmd = Nothing

    ExecuteScalarOnConnection = value
End Function

' Gets a recordset from a SQL query.
' @param sql The SQL query to execute.
' @param params A dictionary of parameters (optional).
' @return An ADODB.Recordset object.
Public Function GetRecordset(ByVal sql As String, Optional ByVal params As Scripting.Dictionary) As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Set conn = GetDbConnection()
    Set rs = GetRecordsetOnConnection(conn, sql, params)

    Set GetRecordset = rs
End Function

' Gets a recordset from a SQL query on an existing connection.
' @param conn The existing ADODB.Connection object.
' @param sql The SQL query to execute.
' @param params A dictionary of parameters (optional).
' @return An ADODB.Recordset object.
Public Function GetRecordsetOnConnection(ByRef conn As ADODB.Connection, ByVal sql As String, Optional ByVal params As Scripting.Dictionary) As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset

    Set cmd = GetDbCommand(conn, sql, params)
    Set rs = cmd.Execute

    Set GetRecordsetOnConnection = rs
End Function

' Creates a new ADODB.Command object with the specified SQL query and parameters.
' @param conn The existing ADODB.Connection object.
' @param sql The SQL query to execute.
' @param params A dictionary of parameters (optional).
' @return An ADODB.Command object.
Public Function GetDbCommand(ByRef conn As ADODB.Connection, ByVal sql As String, Optional ByVal params As Scripting.Dictionary) As ADODB.Command
    Dim cmd As ADODB.Command
    Dim dbParamName As String
    Dim paramName As String
    Dim paramValue As Variant
    Dim paramType As DataTypeEnum
    Dim key As Variant
    Dim extendedSql As String

    'We will declare some parameters so the user can use named parameters
    'like .NET does
    extendedSql = ""

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = CommandTimeout

    If Not params Is Nothing Then
        For Each key In params.Keys
            If Not Left(CStr(key), 1) = "@" Then
                paramName = "@"
            Else
                paramName = ""
            End If
        
            paramName = paramName & key
            dbParamName = paramName & "_db"
            paramValue = params(key)

            'Add the dbParamName to the actual parameter to stop it colliding with the named parameter
            cmd.Parameters.Append GetDbParameter(dbParamName, paramValue)

            extendedSql = extendedSql & "DECLARE " & paramName & " " & GetSqlServerTypeFromDataTypeEnum(GetDataTypeEnumFromVariant(paramValue)) & " = ?;" & vbCrLf
        Next key
    End If

    extendedSql = extendedSql & vbCrLf & sql
    cmd.CommandText = extendedSql

    Set GetDbCommand = cmd
End Function

' Creates a new ADODB.Parameter object with the specified name and value.
' @param name The name of the parameter.
' @param value The value of the parameter.
' @return An ADODB.Parameter object.
Public Function GetDbParameter(ByVal name As String, ByVal value As Variant) As ADODB.Parameter
    Dim param As ADODB.Parameter
    Dim paramType As DataTypeEnum

    If (value = Null) Then
        paramType = adVariant
    Else
        paramType = GetDataTypeEnumFromVariant(value)
    End If

    Set param = New ADODB.Parameter
    param.name = name
    param.Type = paramType
    param.value = value
    
    If paramType = adVarChar Then
        param.Size = 2147483646
    End If

    Set GetDbParameter = param
End Function

' Gets the SQL Server data type from the ADODB data type.
' @param dataType The ADODB data type.
' @return The SQL Server data type as a string.
Private Function GetSqlServerTypeFromDataTypeEnum(ByVal dataType As DataTypeEnum) As String
    Select Case dataType
        Case adInteger
            GetSqlServerTypeFromDataTypeEnum = "INT"
        Case adBigInt
            GetSqlServerTypeFromDataTypeEnum = "BIGINT"
        Case adSingle
            GetSqlServerTypeFromDataTypeEnum = "REAL"
        Case adDouble
            GetSqlServerTypeFromDataTypeEnum = "FLOAT"
        Case adCurrency
            GetSqlServerTypeFromDataTypeEnum = "MONEY"
        Case adDate
            GetSqlServerTypeFromDataTypeEnum = "DATETIMEOFFSET"
        Case adBoolean
            GetSqlServerTypeFromDataTypeEnum = "BIT"
        Case Else
            GetSqlServerTypeFromDataTypeEnum = "VARCHAR(MAX)"
    End Select
End Function

' Gets the ADODB data type from a Variant value.
' @param value The Variant value.
' @return The ADODB data type as a DataTypeEnum.
Private Function GetDataTypeEnumFromVariant(ByVal value As Variant) As DataTypeEnum
    Select Case VarType(value)
        Case vbString
            GetDataTypeEnumFromVariant = adVarChar
        Case vbInteger
            GetDataTypeEnumFromVariant = adInteger
        Case vbLong
            GetDataTypeEnumFromVariant = adBigInt
        Case vbSingle
            GetDataTypeEnumFromVariant = adSingle
        Case vbDouble
            GetDataTypeEnumFromVariant = adDouble
        Case vbCurrency
            GetDataTypeEnumFromVariant = adCurrency
        Case vbDate
            GetDataTypeEnumFromVariant = adDate
        Case vbBoolean
            GetDataTypeEnumFromVariant = adBoolean
        Case Else
            GetDataTypeEnumFromVariant = adVariant
    End Select
End Function

' Protects a string value from SQL injection by escaping single quotes.
' Take great care when using this function as it is not foolproof.
' e.g. it will not protect against double quotes or if the injection into the query
' is not injecting into part of a string
' @param value The string value to protect.
' @return The protected string value.
Public Function SqlInjectionProtect(ByVal value As String) As String
    SqlInjectionProtect = Replace(value, "'", "''")
End Function

' Cleans up a recordset by closing it and its connection.
' @param rs The ADODB.Recordset object to clean up.
' @param closeConnection Whether to close the active connection (optional).
Public Sub CleanUpRecordset(ByRef rs As ADODB.Recordset, Optional ByVal closeConnection As Boolean = True)
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
            rs.Close
        End If

        If rs.ActiveConnection.State = adStateOpen And closeConnection Then
            rs.ActiveConnection.Close
        End If

        Set rs = Nothing
    End If
End Sub
