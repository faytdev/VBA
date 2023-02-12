Attribute VB_Name = "Launch"
Option Base 1
Option Explicit

Const DataSheetName As String = "_DistroManager-DataSheet"

Function MultiNewLines(Optional Amount As Integer = 1) As String
    Dim ReturnVal As String
    Dim i As Integer
    For i = 1 To Amount
        ReturnVal = ReturnVal & vbNewLine
    Next i
    MultiNewLines = ReturnVal
End Function

Function DataSheetCheck() As Boolean
    Dim wSheet As Worksheet
    For Each wSheet In ThisWorkbook.Worksheets
        If wSheet.Name = DataSheetName Then
            DataSheetCheck = True
            Set wSheet = Nothing
            Exit Function
        End If
    Next wSheet
    DataSheetCheck = False
End Function

Function CreateDataSheet() As Boolean
    If MsgBox( _
              "Missing Datasheet." & MultiNewLines(2) & _
              "Do you want to create it?" & MultiNewLines(2), _
              Buttons:=vbYesNo, _
              Title:="DistorManager: Missing DataSheet" _
             ) _
    = vbYes Then
        Dim wSheet As Worksheet
        Set wSheet = ThisWorkbook.Worksheets.Add
        With wSheet
            .Name = DataSheetName
            .Visible = xlSheetVisible 'xlSheetVeryHidden
            .Range("A1").Value = "Group Name"
            .Range("A2").Value = "GroupOne (Change)"
            .Range("B1").Value = "Group Members"
            .Range("B2").Value = "change@email.com;change2@email.com"
        End With
        Set wSheet = Nothing
        CreateDataSheet = True
        Exit Function
    Else
        CreateDataSheet = False
    End If
End Function

Sub LaunchForm()
    If DataSheetCheck = False Then
        If CreateDataSheet = True Then
            DistroManager.Show
        End If
    Else
        DistroManager.Show
    End If
End Sub
