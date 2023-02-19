Attribute VB_Name = "Logger"
Option Explicit

'Name:          Logger
'Description:   Simple Logger Module. Has three log levels {INFO, DEBUG, ERROR}
'               Writes to array until Write method is called.
'Note:          You must flush the Buffer if you intend to do multiple writes. Sees Docs on github
'Created By:    Fayt.Dev { https://github.com/faytdev/VBA/Logger }

Public Enum logLevel
    LogLevelInfo = 0
    LogLevelDebug = 1
    LogLevelError = 2
End Enum

Private pPattern    As String
Private pSymbol     As String
Private Buffer      As Variant

Public Property Get Pattern() As String
    Pattern = pPattern
End Property

Public Property Let Pattern(NewPattern As String)
    pPattern = NewPattern
End Property

Public Property Get Symbol() As String
    Symbol = pSymbol
End Property

Public Property Let Symbol(NewSymbol As String)
    pSymbol = NewSymbol
End Property

Public Property Get BufferArray() As Variant
    BufferArray = Buffer
End Property

Private Sub WriteToBuffer(msg As String)
    If IsArray(Buffer) = False Then
        ReDim Buffer(0 To 0)
    Else
        ReDim Preserve Buffer(0 To UBound(Buffer) + 1)
    End If
    Buffer(UBound(Buffer, 1)) = msg
End Sub

Private Sub SetDefaults()
    If Logger.Pattern = "False" Or Logger.Pattern = vbNullString Then Logger.Pattern = "< " & Format$(Date + Time, "mm/dd/yyyy@hh:mm:ss") & " > ::#::  "
    If Logger.Symbol = "False" Or Logger.Symbol = vbNullString Then Logger.Symbol = "#"
End Sub

Public Sub Flush()
    Erase Buffer
    Buffer = 0
End Sub

Public Sub FlushPatternAndSymbol()
    Logger.Pattern = vbNullString
    Logger.Symbol = vbNullString
End Sub

Public Sub Log(message As String, Optional lvl As logLevel = logLevel.LogLevelInfo)
    Select Case lvl
        Case logLevel.LogLevelInfo: Logger.LogInfo message
        Case logLevel.LogLevelDebug: Logger.LogDebug message
        Case logLevel.LogLevelError: Logger.LogError message
    End Select
End Sub

Public Sub LogInfo(msg As String)
    SetDefaults
    If InStr(1, Logger.Pattern, Logger.Symbol, vbTextCompare) > 0 Then
        WriteToBuffer Replace(Logger.Pattern, Logger.Symbol, "INFO ") & msg
    Else
        WriteToBuffer Logger.Pattern & "INFO" & msg
    End If
End Sub

Public Sub LogDebug(msg As String)
    SetDefaults
    If InStr(1, Logger.Pattern, Logger.Symbol, vbTextCompare) > 0 Then
        WriteToBuffer Replace(Logger.Pattern, Logger.Symbol, "DEBUG") & msg
    Else
        WriteToBuffer Logger.Pattern & "DEBUG " & msg
    End If
End Sub

Public Sub LogError(msg As String)
    SetDefaults
    If InStr(1, Logger.Pattern, Logger.Symbol, vbTextCompare) > 0 Then
        WriteToBuffer Replace(Logger.Pattern, Logger.Symbol, "ERROR") & msg
    Else
        WriteToBuffer Logger.Pattern & "ERROR " & msg
    End If
End Sub

Public Sub WriteToFile(FilePath As String, Optional FlushBuffer As Boolean = False, Optional FlushPatternSymbol As Boolean = False)
    Dim Fso As New FileSystemObject
    Dim f As TextStream
    Set f = Fso.OpenTextFile(FilePath, ForAppending, True)
    f.Write Join(Buffer, vbNewLine) & vbNewLine
    f.Close
    Set f = Nothing
    Set Fso = Nothing
    If FlushBuffer Then Logger.Flush
    If FlushPatternSymbol Then Logger.FlushPatternAndSymbol
End Sub

Public Sub WriteToConsole(Optional FlushBuffer As Boolean = False, Optional FlushPatternSymbol As Boolean = False)
    Debug.Print Join(Buffer, vbNewLine) & vbNewLine
    If FlushBuffer Then Logger.Flush
    If FlushPatternSymbol Then Logger.FlushPatternAndSymbol
End Sub

Public Sub WriteToRange(Target As Range, Optional FlushBuffer As Boolean = False, Optional FlushPatternSymbol As Boolean = False)
    Dim WriteBuffer As Variant
    ReDim WriteBuffer(0 To UBound(Buffer, 1), 0 To 0)
    Dim i As Long
    For i = 0 To UBound(WriteBuffer, 1)
        WriteBuffer(i, 0) = Buffer(i)
    Next i
    Set Target = Target.Cells(1, 1)
    Target.Resize(UBound(WriteBuffer, 1) + 1, 1).Value = WriteBuffer
    Erase WriteBuffer
    Set Target = Nothing
    If FlushBuffer Then Logger.Flush
    If FlushPatternSymbol Then Logger.FlushPatternAndSymbol
End Sub
