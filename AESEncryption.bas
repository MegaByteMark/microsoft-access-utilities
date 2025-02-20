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

' Add references to Microsoft XML, v6.0 and Microsoft Scripting Runtime

Public Function Encrypt(ByVal plainText As String, ByVal key As String) As String
    Dim aes As Object
    Dim iv As String
    Dim encryptor As Object
    Dim plainBytes() As Byte
    Dim cipherBytes() As Byte

    Set aes = CreateObject("System.Security.Cryptography.AesManaged")
    
    aes.KeySize = 256
    aes.Key = StrConv(key, vbFromUnicode)

    'Generate a new unique IV for each encryption
    aes.GenerateIV
    
    iv = StrConv(aes.IV, vbUnicode)

    Set encryptor = aes.CreateEncryptor()
    plainBytes = StrConv(plainText, vbFromUnicode)
    cipherBytes = encryptor.TransformFinalBlock(plainBytes, 0, UBound(plainBytes) + 1)
    
    'Append the IV to the start of the cipher text to make it unique and enable us to extract the 
    'IV in the decrypt function
    Encrypt = iv & StrConv(cipherBytes, vbUnicode)
End Function

Public Function Decrypt(ByVal cipherText As String, ByVal key As String) As String
    Dim aes As Object
    Dim iv As String
    Dim cipherBytes() As Byte
    Dim decryptor As Object
    Dim plainBytes() As Byte

    Set aes = CreateObject("System.Security.Cryptography.AesManaged")

    aes.KeySize = 256
    aes.Key = StrConv(key, vbFromUnicode)

    'Extract the IV from the start of the cipher text
    iv = Left(cipherText, 16)
    aes.IV = StrConv(iv, vbFromUnicode)
    cipherBytes = StrConv(Mid(cipherText, 17), vbFromUnicode)

    Set decryptor = aes.CreateDecryptor()
    plainBytes = decryptor.TransformFinalBlock(cipherBytes, 0, UBound(cipherBytes) + 1)
    
    Decrypt = StrConv(plainBytes, vbUnicode)
End Function