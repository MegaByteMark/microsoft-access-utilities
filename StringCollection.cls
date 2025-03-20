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

Private col As Collection

Public Sub Class_Initialize()
    Set col = New Collection
End Sub

Public Sub Add(ByVal InputString As String)

    col.Add InputString

End Sub

Public Function Count() As Integer

    Count = col.Count

End Function

Public Function Item(ByVal Index As Integer) As String

    Item = col.Item(Index)

End Function

Public Sub Remove(ByVal SearchString As String)

    col.Remove IndexOf(SearchString)

End Sub

Public Function Contains(ByVal SearchString As String, Optional ByVal IgnoreCase As Boolean = False) As Boolean

    Dim Index As Integer
    Dim found As Boolean
    
    Index = IndexOf(SearchString, IgnoreCase)
    
    found = Index > -1
    
    Contains = found

End Function

Public Function IndexOf(ByVal SearchString As String, Optional ByVal IgnoreCase As Boolean = False) As Integer

    Dim Index As Integer
    Dim found As Boolean
    Dim Item As String

    Index = 1
    found = False
    
    If IgnoreCase Then
        
        SearchString = LCase(SearchString)
        
    End If
    
    While Not found And Index <= col.Count
    
        Item = col.Item(Index)
        
        If IgnoreCase Then
        
            Item = LCase(Item)
        
        End If
        
        If Item = SearchString Then
            
            found = True
            
        Else
            
            Index = Index + 1
            
        End If
    
    Wend
    
    If Not found Then
    
        Index = -1
    
    End If

    IndexOf = Index

End Function
