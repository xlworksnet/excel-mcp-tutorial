Attribute VB_Name = "ScheduleModule"
' 진척률을 다시 계산하고, 완료된 업무는 회색, 지연된 업무는 빨간색으로 표시한다.
' 결과는 "완료 n건, 지연 n건" 문자열로 돌려준다.
Public Function UpdateSchedule() As String
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long, nDone As Long, nDelay As Long

    Set ws = ThisWorkbook.Worksheets("일정")
    Application.Calculate
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    For r = 2 To lastRow
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Interior
            If ws.Cells(r, 7).Value >= 1 Then
                .Color = RGB(217, 217, 217)
                nDone = nDone + 1
            ElseIf ws.Cells(r, 8).Value = "지연" Then
                .Color = RGB(255, 199, 206)
                nDelay = nDelay + 1
            Else
                .ColorIndex = xlNone
            End If
        End With
    Next r

    UpdateSchedule = "완료 " & nDone & "건, 지연 " & nDelay & "건"
End Function
