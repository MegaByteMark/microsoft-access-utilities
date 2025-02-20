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

Option Compare Database
Option Explicit

Private Declare PtrSafe Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" ( _
    ByVal lpszUsername As String, _
    ByVal lpszDomain As String, _
    ByVal lpszPassword As String, _
    ByVal dwLogonType As Long, _
    ByVal dwLogonProvider As Long, _
    phToken As Long) As Long

Private Declare PtrSafe Function ImpersonateLoggedOnUser Lib "advapi32.dll" ( _
    ByVal hToken As Long) As Long

Private Declare PtrSafe Function RevertToSelf Lib "advapi32.dll" () As Long

Private Declare PtrSafe Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long) As Long
    
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
ByVal lpBuffer As String, _
ByRef nSize As Long) As Long

Public Function ImpersonateUser(username As String, domain As String, password As String) As Long
    Dim hToken As Long
    Dim result As Long
    
    'dwLogonType = 2 (LOGON32_LOGON_INTERACTIVE)
    'dwLogonProvider = 0 (LOGON32_PROVIDER_DEFAULT)
    result = LogonUser(username, domain, password, 2, 0, hToken)
    
    If result = 0 Then
        Err.Raise vbObjectError + 1, "ImpersonateUser", "LogonUser failed with error: " & Err.LastDllError
        Exit Function
    End If
    
    result = ImpersonateLoggedOnUser(hToken)

    If result = 0 Then
        Err.Raise vbObjectError + 2, "ImpersonateUser", "ImpersonateLoggedOnUser failed with error: " & Err.LastDllError
        CloseHandle hToken
        Exit Function
    End If

    ImpersonateUser = hToken

End Function

Public Sub Revert(hToken As Long)
    Dim result As Long
    
    result = RevertToSelf()
    
    If result = 0 Then
        Err.Raise vbObjectError + 1, "Revert", "RevertToSelf failed with error: " & Err.LastDllError
        Exit Sub
    End If

    CloseHandle hToken

    If result = 0 Then
        Err.Raise vbObjectError + 2, "Revert", "CloseHandle failed with error: " & Err.LastDllError
        Exit Sub
    End If
End Sub

Public Function GetCurrentContextUserName() As String
    Dim buffer As String
    Dim bufferSize As Long
    Dim result As Long
    
    bufferSize = 255
    buffer = String(bufferSize, vbNullChar)
    
    result = GetUserName(buffer, bufferSize)
    
    If result <> 0 Then
        GetCurrentContextUserName = Left(buffer, bufferSize - 1)
    Else
        GetCurrentContextUserName = "Unknown User"
    End If
End Function
