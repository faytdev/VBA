VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DistroManager 
   Caption         =   "Distro Manager"
   ClientHeight    =   7740
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6490
   OleObjectBlob   =   "DistroManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DistroManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit

Const DataSheetName As String = "_DistroManager-DataSheet"

Private Enum dmData
    Groups = 1
    members = 2
End Enum

Private Enum dmEditMode
    View = 0
    Edit = 1
    Add = 2
End Enum

Private Sub cboxEditMode_Change()
    Select Case Me.cboxEditMode.ListIndex
        Case dmEditMode.View:
            Me.tboxMemberValue.Enabled = False
            Me.cbtnRemove.Enabled = False
            With Me.cbtnAddEdit
                .Caption = "View"
                .Enabled = False
            End With
        Case dmEditMode.Edit:
            Me.tboxMemberValue.Enabled = True
            Me.cbtnRemove.Enabled = True
            With Me.cbtnAddEdit
                .Caption = "Edit"
                .Enabled = True
            End With
        Case dmEditMode.Add:
            Me.tboxMemberValue.Enabled = True
            Me.cbtnRemove.Enabled = True
            With Me.cbtnAddEdit
                .Caption = "Add"
                .Enabled = True
            End With
        Case Else:
            Me.tboxMemberValue.Enabled = True
    End Select
End Sub

Private Sub cboxGroupSelect_Change()
    Me.lboxGroupMembers.Clear
    Dim SplitMembers As Variant
    SplitMembers = GetElement(GetDistroData(), members)
    Dim GetNext As Variant
    If IsEmpty(SplitMembers) Then
        Me.lboxGroupMembers.AddItem vbNullString
    Else
        For Each GetNext In SplitMembers
            Me.lboxGroupMembers.AddItem CStr(GetNext)
        Next GetNext
        Erase SplitMembers
    End If
End Sub

Private Sub cbtnAddEdit_Click()
    If Me.tboxMemberValue.Text = vbNullString Then Exit Sub
    Dim LastRow As Long
    LastRow = ThisWorkbook.Worksheets(DataSheetName).Range("A" & Rows.count).End(xlUp).Row + 1
    Select Case Me.cboxEditMode.ListIndex
        Case dmEditMode.Add:
            If Me.rbtnGroups.Value = True Then
                ThisWorkbook.Worksheets(DataSheetName).Range("A" & LastRow).Value = Me.tboxMemberValue.Text
                With Me.lboxGroupMembers
                    .Clear
                    .List = GetElement(GetDistroData(), Groups)
                End With
                With Me.lboxGroupMembers
                    .Clear
                    .List = GetElement(GetDistroData(), Groups)
                End With
            ElseIf Me.rbtnMembers.Value = True Then
'                ThisWorkbook.Worksheets(DataSheetName).Range("B" & Me.cboxGroupSelect.ListIndex + 2).Value = _
'                ThisWorkbook.Worksheets(DataSheetName).Range("B" & Me.cboxGroupSelect.ListIndex + 2).Value & ";" & Me.tboxMemberValue.Text
                With ThisWorkbook.Worksheets(DataSheetName).Range("B" & Me.cboxGroupSelect.ListIndex + 2)
                    If .Value = vbNullString Then
                        .Value = Me.tboxMemberValue.Text
                    Else
                        .Value = .Value & ";" & Me.tboxMemberValue.Text
                    End If
                End With
                With Me.lboxGroupMembers
                    .Clear
                    .List = GetElement(GetDistroData(), Groups)
                End With
                With Me.lboxGroupMembers
                    .Clear
                    .List = GetElement(GetDistroData(), members)
                End With
            End If
        Case dmEditMode.Edit:
            If Me.rbtnGroups.Value = True Then
                If Me.lboxGroupMembers.ListIndex < 0 Then
                    MsgBox "Select Item to Edit", vbOKOnly, "Distro Manager: Select Item"
                    Exit Sub
                End If
                ThisWorkbook.Worksheets(DataSheetName).Range("A" & Me.lboxGroupMembers.ListIndex + 2).Value = Me.tboxMemberValue.Text
                With Me.lboxGroupMembers
                    .Clear
                    .List = GetElement(GetDistroData(), Groups)
                End With
                With Me.lboxGroupMembers
                    .Clear
                    .List = GetElement(GetDistroData(), Groups)
                End With
            ElseIf Me.rbtnMembers.Value = True Then
                Dim arr As Variant
                arr = Split(ThisWorkbook.Worksheets(DataSheetName).Range("B" & Me.cboxGroupSelect.ListIndex + 2).Value, ";")
                arr(Me.lboxGroupMembers.ListIndex) = Me.tboxMemberValue.Text
                ThisWorkbook.Worksheets(DataSheetName).Range("B" & Me.cboxGroupSelect.ListIndex + 2).Value = Join(arr, ";")
                Erase arr
                With Me.lboxGroupMembers
                    .Clear
                    .List = GetElement(GetDistroData(), Groups)
                End With
                With Me.lboxGroupMembers
                    .Clear
                    .List = GetElement(GetDistroData(), members)
                End With
            End If
    End Select
End Sub

Private Sub cbtnMakeNamedRanges_Click()
    Dim CheckName As Name
    Dim GroupArray As Variant
    Dim MissingArray As Variant
    GroupArray = GetDistroData(True)
    Dim ArrOffset As Integer
    ArrOffset = 2
    Dim i As Integer
    Dim InNames As Boolean
    InNames = False
    If ThisWorkbook.Names.count <> 0 Then
        For Each CheckName In ThisWorkbook.Names
            For i = LBound(GroupArray, 1) To UBound(GroupArray, 1)
                If CheckName.Name = GroupArray(i, 1) Then
                    InNames = True
                    Exit For
                End If
            Next i
            If InNames = False Then
                If IsEmpty(MissingArray) Then
                    ReDim MissingArray(1 To 1, 1 To 2)
                Else
                    ReDim MissingArray(1 To UBound(MissingArray, 1) + 1, 1 To 2)
                End If
                MissingArray(UBound(MissingArray, 1), 1) = GroupArray(UBound(MissingArray, 1), 1)
                MissingArray(UBound(MissingArray, 1), 2) = GroupArray(UBound(MissingArray, 1), 3)
            End If
            InNames = False
        Next CheckName
        If IsEmpty(MissingArray) Then
            InNames = True
        End If
    End If
    If IsEmpty(MissingArray) Then
            MissingArray = GetDistroData(True)
            ArrOffset = 3
    End If
    If InNames = False Then
        For i = LBound(MissingArray, 1) To UBound(MissingArray, 1)
            ThisWorkbook.Names.Add Name:=MissingArray(i, 1), RefersToR1C1:="='" & DataSheetName & "'!" & MissingArray(i, ArrOffset)
            ThisWorkbook.Names(MissingArray(i, 1)).Comment = MissingArray(i, ArrOffset)
        Next i
    Else
        MsgBox "All groups already in name manger.", vbOKOnly, "DistroManager: Names Already Exist"
    End If
End Sub

Private Sub cbtnRemove_Click()
    If Me.lboxGroupMembers.ListIndex >= 0 Then
        If Me.rbtnGroups Then
            ThisWorkbook.Worksheets(DataSheetName).Range("A" & Me.lboxGroupMembers.ListIndex + 2).EntireRow.Delete Shift:=xlShiftUp
            With Me.lboxGroupMembers
                .Clear
                .List = GetElement(GetDistroData(), Groups)
            End With
            With Me.lboxGroupMembers
                .Clear
                .List = GetElement(GetDistroData(), Groups)
            End With
        End If
        If Me.rbtnMembers Then
            Dim SplitArray As Variant
            Dim WriteArray As Variant
            SplitArray = GetElement(GetDistroData(), members)
            ReDim WriteArray(0 To UBound(SplitArray))
            Dim i As Integer
            Dim count As Integer
            count = 0
            For i = LBound(SplitArray, 1) To UBound(SplitArray, 1)
                If i <> Me.lboxGroupMembers.ListIndex Then
                    WriteArray(count) = SplitArray(i)
                    count = count + 1
                End If
            Next i
            ThisWorkbook.Worksheets(DataSheetName).Range("B" & Me.cboxGroupSelect.ListIndex + 2).Value = Join(WriteArray, ";")
            Erase SplitArray
            Erase WriteArray
            With Me.lboxGroupMembers
                .Clear
                .List = GetElement(GetDistroData(), Groups)
            End With
            With Me.lboxGroupMembers
                .Clear
                .List = GetElement(GetDistroData(), members)
            End With
        End If
    Else
        MsgBox "Select item to delete first.", vbOKOnly, "Distro Manager: Select Item To Remove"
    End If
End Sub

Private Sub chboxAdvanceOptions_Change()
    Select Case chboxAdvanceOptions.Value
        Case True:
           Me.frmAdvanceOptions.Visible = True
           Me.frmManage.Top = 96
           Me.Height = 415
        Case False:
            Me.frmAdvanceOptions.Visible = False
            Me.frmManage.Top = 54
            Me.Height = 385
    End Select
End Sub

Private Sub lboxGroupMembers_Change()
    If Me.lboxGroupMembers.ListIndex >= 0 And _
       Me.cboxEditMode.ListIndex <> dmEditMode.View And _
       Me.cboxEditMode.ListIndex <> dmEditMode.Add _
    Then
        Me.tboxMemberValue.Text = Me.lboxGroupMembers.List(Me.lboxGroupMembers.ListIndex)
    Else
        Me.tboxMemberValue.Text = vbNullString
    End If
End Sub

Private Sub rbtnGroups_Click()
    Me.rbtnMembers.Value = False
    Me.cboxGroupSelect.Enabled = False
    With Me.lboxGroupMembers
        .Clear
        .List = GetElement(GetDistroData(), Groups)
    End With
End Sub

Private Sub rbtnMembers_Click()
    Me.rbtnGroups.Value = False
    With Me.cboxGroupSelect
        .Enabled = True
        .Clear
        .List = GetElement(GetDistroData(), Groups)
        If .ListCount > 1 Then
            .Value = .List(.ListCount - 1)
            .Value = .List(0)
        Else
            .AddItem "#CHANGE:ITEM#"
            .Value = .List(.ListCount - 1)
            .Value = .List(0)
            .RemoveItem .ListCount - 1
        End If
    End With
End Sub

Private Function GetElement(ByVal arr As Variant, Val As dmData) As Variant
    Dim ReturnArr As Variant
    Dim i As Integer
    If Val = Groups Then
        ReDim ReturnArr(1 To UBound(arr, 1))
        For i = LBound(arr, 1) To UBound(arr, 1)
            ReturnArr(i) = arr(i, Val)
        Next i
        GetElement = ReturnArr
        Erase ReturnArr
    ElseIf Val = members Then
        If Me.cboxGroupSelect.ListIndex >= 0 Then
            If UBound(arr, 1) = 1 Then
                GetElement = Split(CStr(arr(1, 2)), ";")
            Else
                For i = LBound(arr, 1) To UBound(arr, 1)
                    If arr(i, 1) = Me.cboxGroupSelect.List(Me.cboxGroupSelect.ListIndex) Then
                        GetElement = Split(CStr(arr(i, 2)), ";")
                        Exit For
                    End If
                Next i
            End If
        Else
            GetElement = Split(CStr(arr(1, 2)), ";")
        End If
    End If
    Erase arr
End Function

Private Function GetDistroData(Optional GetAddress As Boolean = False) As Variant
    Dim DistroArray As Variant
    Dim LastRow As Long
    LastRow = ThisWorkbook.Worksheets(DataSheetName).Range("A" & ThisWorkbook.Worksheets(DataSheetName).Rows.count).End(xlUp).Row
    DistroArray = ThisWorkbook.Worksheets(DataSheetName).Range("A2:B" & LastRow)
    If GetAddress = False Then
        GetDistroData = DistroArray
    Else
        ReDim Preserve DistroArray(LBound(DistroArray, 1) To UBound(DistroArray, 1), 1 To 3)
        Dim Cell As Range
        For Each Cell In ThisWorkbook.Worksheets(DataSheetName).Range("A2:A" & LastRow)
            DistroArray(Cell.Row - 1, 3) = Cell.Offset(0, 1).Address(ReferenceStyle:=xlR1C1)
        Next Cell
        Set Cell = Nothing
        GetDistroData = DistroArray
    End If
    Erase DistroArray
End Function

Private Sub UserForm_Initialize()
    With Me.cboxGroupSelect
        .List = GetElement(GetDistroData(), Groups)
        .Value = .List(0)
    End With
    With Me.cboxEditMode
        .AddItem "View"
        .AddItem "Edit"
        .AddItem "Add"
        .Value = .List(0)
    End With
    Me.rbtnMembers.Value = True
    Me.frmAdvanceOptions.Visible = False
    Me.frmManage.Top = 54
    Me.Height = 410
    Me.cbtnRemove.Enabled = False
End Sub
