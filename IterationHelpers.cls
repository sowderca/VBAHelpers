VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IterationHelpers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub ProcessWorkSheet()
    ProcessEachRow "Sheet2", "A:A"
End Sub

Public Function ProcessEachRow(Optional sheetName As String, Optional columnName As String)
    With Worksheets(sheetName)
        For i = 1 To .Rows.Count
            cellValue = .Range(columnName).Cells(i, 1).Value
            If cellValue <> "" Then
                FilterOnMatch (cellValue)
            Else
                If IsEmpty(cellValue) = True Then Exit For
            End If
        Next i
    End With
End Function

Public Function FilterOnMatch(Optional entryMatch As String)
    With Worksheets("Sheet1")
        .Range("A:A").AutoFilter Field:=1, Criteria1:=entryMatch
    End With
End Function

