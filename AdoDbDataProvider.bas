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
Private connTimeout As Integer = 0
Private cmdTimeout As Integer = 0

Public Property Let ConnectionString(value As String)
    connString = value
End Property

Public Property Get ConnectionString() As String
    ConnectionString = connString
End Property

Public Property Let ConnectionTimeout(value As Integer)
    connTimeout = value
End Property

Public Property Get ConnectionTimeout() As String
    ConnectionTimeout = connTimeout
End Property

Public Property Let CommandTimeout(value As Integer)
    cmdTimeout = value
End Property

Public Property Get CommandTimeout() As String
    CommandTimeout = cmdTimeout
End Property

Public Sub Initialize(ByVal connString As String, Optional ByVal connTimeout As Integer = 0, Optional ByVal cmdTimeout As Integer = 0)
    ConnectionString = connString
    ConnectionTimeout = connTimeout
    CommandTimeout = cmdTimeout
End Sub

Public Function GetDbConnection(Optional ByVal connectionTimeout As Integer = connTimeout) As ADODB.Connection
    Dim conn As ADODB.Connection

    If ConnectionString = "" Then
        Err.Raise 1, "GetDbConnection", "Connection string not set"
    End If

    Set conn = New ADODB.Connection
    conn.ConnectionString = ConnectionString
    conn.ConnectionTimeout = connectionTimeout
    
    conn.Open
    
    Set GetConnection = conn
End Function

Public Function ExecuteNonQuery(ByVal sql As String, Optional ByVal params As Variant) As Long
    Dim conn As ADODB.Connection
    Dim affectedRows As Long

    Set conn = GetDbConnection()
    affectedRows = ExecuteNonQueryOnConnection(conn, sql, params)
    conn.Close

    Set conn = Nothing

    ExecuteNonQuery = affectedRows
End Function

Public Function ExecuteNonQueryOnConnection(ByRef conn As ADODB.Connection, ByVal sql As String, Optional ByVal params As Variant) As Long
    Dim cmd As ADODB.Command
    Dim affectedRows As Long

    Set cmd = GetDbCommand(conn, sql, params)

    cmd.Execute affectedRows

    Set cmd = Nothing

    ExecuteNonQuery = affectedRows
End Function

Public Function ExecuteScalar(ByVal sql As String, Optional ByVal params As Variant) As Variant
    Dim conn As ADODB.Connection
    Dim value As Variant

    Set conn = GetDbConnection()
    value = ExecuteScalarOnConnection(conn, sql, params)
    conn.Close

    Set conn = Nothing

    ExecuteScalar = value
End Function

Public Function ExecuteScalarOnConnection(ByRef conn As ADODB.Connection, ByVal sql As String, Optional ByVal params As Variant) As Variant
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim value As Variant

    Set cmd = GetDbCommand(conn, sql, params)

    Set rs = cmd.Execute

    If Not rs.EOF Then
        value = rs.Fields(0).Value
    End If

    rs.Close
    Set rs = Nothing
    Set cmd = Nothing

    ExecuteScalar = value
End Function

Public Function GetRecordset(ByVal sql As String, Optional ByVal params As Variant) As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Set conn = GetDbConnection()
    Set rs = GetRecordsetOnConnection(conn, sql, params)
    conn.Close

    Set conn = Nothing

    Set GetRecordset = rs
End Function

Public Function GetRecordsetOnConnection(ByRef conn As ADODB.Connection, ByVal sql As String, Optional ByVal params As Scripting.Dictionary) As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset

    Set cmd = GetDbCommand(conn, sql, params)
    Set rs = cmd.Execute
    Set cmd = Nothing

    Set GetRecordsetOnConnection = rs
End Function

Public Function GetDbCommand(ByRef conn As ADODB.Connection, ByVal sql As String, Optional ByVal params As Scripting.Dictionary) As ADODB.Command
    Dim cmd As ADODB.Command
    Dim paramName As String
    Dim paramValue As Variant
    Dim paramType As DataTypeEnum

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = CommandTimeout

    If Not params Is Nothing Then
        For Each key In dict.Keys
            paramName = "@" & key
            paramValue = dict(key)

            cmd.Parameters.Append GetDbParameter(paramName, paramValue)
        Next key
    End If

    Set GetDbCommand = cmd
End Function

Public Function GetDbParameter (ByVal name As String, ByVal value As Variant) As ADODB.Parameter
    Dim param As ADODB.Parameter
    Dim paramType As DataTypeEnum

    if(value = Null) Then
        paramType = adVariant
    Else
        paramType = GetDataTypeEnumFromVariant(value)
    End If

    Set param = New ADODB.Parameter
    param.Name = name
    param.Type = paramType
    param.Value = value

    Set GetDbParameter = param
End Function

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

Public Function SqlInjectionProtect(ByVal value As String) As String
    SqlInjectionProtect = Replace(value, "'", "''")
End Function