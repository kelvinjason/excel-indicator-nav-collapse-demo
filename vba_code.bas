Attribute VB_Name = "IndicatorToggle"
Option Explicit

Public Sub ToggleCategory(ByVal categoryRow As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Indicators")

    Dim startRow As Long
    Dim endRow As Long
    startRow = CLng(ws.Cells(categoryRow, "N").Value)
    endRow = CLng(ws.Cells(categoryRow, "O").Value)

    If startRow = 0 Or endRow = 0 Then Exit Sub

    Dim wasExpanded As Boolean
    wasExpanded = Not ws.Rows(startRow & ":" & endRow).Hidden

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    Application.ScreenUpdating = False
    Dim r As Long
    Dim sRow As Long
    Dim eRow As Long

    For r = 4 To lastRow
        If Left$(CStr(ws.Cells(r, "C").Value), 4) = "CAT-" Then
            sRow = Val(ws.Cells(r, "N").Value)
            eRow = Val(ws.Cells(r, "O").Value)
            If sRow > 0 And eRow >= sRow Then
                ws.Rows(sRow & ":" & eRow).Hidden = True
                ws.Cells(r, "A").Value = "+"
            End If
        End If
    Next r

    If wasExpanded Then
        ws.Rows(startRow & ":" & endRow).Hidden = True
        ws.Cells(categoryRow, "A").Value = "+"
    Else
        ws.Rows(startRow & ":" & endRow).Hidden = False
        ws.Cells(categoryRow, "A").Value = "-"
    End If

    Application.Goto ws.Cells(categoryRow, "B"), True
    Application.ScreenUpdating = True
End Sub

Public Sub ExpandCategory1()
    ToggleCategory 4
End Sub

Public Sub ExpandCategory2()
    ToggleCategory 10
End Sub

Public Sub ExpandCategory3()
    ToggleCategory 16
End Sub

Public Sub ExpandCategory4()
    ToggleCategory 22
End Sub
