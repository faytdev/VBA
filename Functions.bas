'########################################################################################################
' Name:         Get Last Occurance
'Description:   Returns postition of the last occurance of search charactor
'Usage:         Int = GetLatOccurance("SEARCH\Value", "\")
'Note:          If GetLastOccurance = 0 Then Findvalue was not found
'Created By:    Fayt.Dev { https://github.com/faytdev/VBA }
'
Function GetLastOccurance(SearchValue As String, FindValue As String) As Integer
    'Created By: Fayt.Dev { https://github.com/faytdev/VBA }
    Dim i As Integer
    For i = Len(SearchValue) To 1 Step -1
        If Mid(SearchValue, i, 1) = FindValue Then
            GetLastOccurance = i
            Exit For
        End If
    Next i
End Function
'
'########################################################################################################

'########################################################################################################
' Name:         GetFileName
'Description:   Returns File name form path string
'               Set SlashChar to change the search charactor (Default = \)
'               Set NotFoundString to check the not found return value (Default = #NF#)
'Requires:      GetLastOccurance or Similar Function/Code
'Usage: Str = GetFilename("SEARCH\Value")
'Created By:    Fayt.Dev { https://github.com/faytdev/VBA }
'
Function GetFileName(FilePath As String, Optional SlashChar As String = "\", Optional NotFoundString As String = "#NF#") As String
    'Created By: Fayt.Dev { https://github.com/faytdev/VBA }
    Dim LastPos As Integer
    LastPos = GetLastOccurance(FilePath, SlashChar)
    If LastPos > 0 Then
        GetFileName = Mid(FilePath, LastPos + 1, (Len(FilePath) - LastPos))
    Else
        GetFileName = NotFoundString
    End If
End Function
'
'########################################################################################################