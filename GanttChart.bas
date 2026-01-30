Attribute VB_Name = "GanttChart"
'===============================================================================
' Visual Gantt - Excel VBA Timeline Generator
' Creates professional Gantt chart with draggable/resizable shapes
'===============================================================================

Option Explicit

' Constants
Private Const DATA_SHEET_NAME As String = "Project Data"
Private Const TIMELINE_SHEET_NAME As String = "Timeline View"
Private Const HEADER_ROWS As Integer = 3
Private Const LABEL_COLS As Integer = 4
Private Const WEEK_COL_WIDTH As Double = 11
Private Const TASK_ROW_HEIGHT As Double = 36
Private Const BAR_HEIGHT As Double = 22
Private Const BAR_TOP_MARGIN As Double = 7

' Column indices in Project Data (1-based)
Private Const COL_TASK_ID As Integer = 1
Private Const COL_TASK_NAME As Integer = 2
Private Const COL_START_DATE As Integer = 3
Private Const COL_END_DATE As Integer = 4
Private Const COL_OWNER As Integer = 6
Private Const COL_PERCENT As Integer = 7
Private Const COL_PROJECT As Integer = 14
Private Const COL_ORIG_END As Integer = 16

' Colors for alternating rows
Private Const ROW_COLOR_1 As Long = 16777215  ' White
Private Const ROW_COLOR_2 As Long = 15921906  ' Light gray #F2F2F2

' Dynamic project color tracking
Private projectColors As Collection
Private colorIndex As Integer
Private colorPalette(0 To 9) As Long

'===============================================================================
' Main Entry Point
'===============================================================================
Public Sub GenerateTimeline()
    Dim wsData As Worksheet
    Dim wsTimeline As Worksheet
    Dim startDate As Date
    Dim endDate As Date
    Dim totalWeeks As Integer
    Dim lastRow As Long
    Dim barsCreated As Integer

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Initialize dynamic color system
    InitializeColors

    Set wsData = GetDataSheet()
    If wsData Is Nothing Then
        MsgBox "Could not find '" & DATA_SHEET_NAME & "' sheet.", vbExclamation
        GoTo Cleanup
    End If

    Set wsTimeline = GetOrCreateTimelineSheet()

    ' Date range: -1 month to +3 months
    startDate = DateSerial(Year(Date), Month(Date) - 1, 1)
    endDate = DateSerial(Year(Date), Month(Date) + 3, 0)
    totalWeeks = Int((endDate - startDate) / 7) + 1

    ' Clear existing
    wsTimeline.Cells.Clear
    ClearAllShapes wsTimeline

    ' Hide gridlines
    wsTimeline.Activate
    ActiveWindow.DisplayGridlines = False

    ' Setup
    SetupTimelineLayout wsTimeline, totalWeeks
    BuildHeaders wsTimeline, startDate, endDate, totalWeeks

    lastRow = wsData.Cells(wsData.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No tasks found.", vbInformation
        GoTo Cleanup
    End If

    barsCreated = BuildTaskBars(wsData, wsTimeline, startDate, endDate, totalWeeks, lastRow)
    AddMonthDividers wsTimeline, startDate, totalWeeks, lastRow - 1
    AddTodayMarker wsTimeline, startDate, totalWeeks, lastRow - 1
    AddChartBorder wsTimeline, lastRow - 1, totalWeeks

    ' Freeze panes
    wsTimeline.Range("E4").Select
    ActiveWindow.FreezePanes = True

    MsgBox "Timeline Generated!" & vbCrLf & vbCrLf & _
           "Task bars created: " & barsCreated & " of " & (lastRow - 1), vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

'===============================================================================
Private Function GetDataSheet() As Worksheet
    On Error Resume Next
    Set GetDataSheet = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    On Error GoTo 0
End Function

'===============================================================================
Private Function GetOrCreateTimelineSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TIMELINE_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(1))
        ws.Name = TIMELINE_SHEET_NAME
    End If
    Set GetOrCreateTimelineSheet = ws
End Function

'===============================================================================
Private Sub ClearAllShapes(ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    On Error GoTo 0
End Sub

'===============================================================================
Private Sub SetupTimelineLayout(ws As Worksheet, totalWeeks As Integer)
    Dim i As Integer

    ' White background everywhere
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    ' Column widths
    ws.Columns(1).ColumnWidth = 15  ' Project
    ws.Columns(2).ColumnWidth = 25  ' Task Name
    ws.Columns(3).ColumnWidth = 6   ' %
    ws.Columns(4).ColumnWidth = 15  ' Owner

    For i = 1 To totalWeeks
        ws.Columns(LABEL_COLS + i).ColumnWidth = WEEK_COL_WIDTH
    Next i

    ' Row heights
    ws.Rows(1).RowHeight = 32
    ws.Rows(2).RowHeight = 24
    ws.Rows(3).RowHeight = 22
End Sub

'===============================================================================
Private Sub BuildHeaders(ws As Worksheet, startDate As Date, endDate As Date, totalWeeks As Integer)
    Dim i As Integer
    Dim weekStart As Date
    Dim fridayDate As Date
    Dim currentMonth As String
    Dim monthStartCol As Integer
    Dim prevMonth As String
    Dim monthBoundaries() As Integer
    Dim boundaryCount As Integer

    ReDim monthBoundaries(0 To totalWeeks)
    boundaryCount = 0

    ' Title - merged above timeline columns only (starting at column D)
    ws.Range(ws.Cells(1, LABEL_COLS + 1), ws.Cells(1, LABEL_COLS + totalWeeks)).Merge
    ws.Cells(1, LABEL_COLS + 1).Value = "Project Timeline: " & Format(startDate, "mmmm d, yyyy") & " - " & Format(endDate, "mmmm d, yyyy")
    ws.Cells(1, LABEL_COLS + 1).Font.Size = 14
    ws.Cells(1, LABEL_COLS + 1).Font.Bold = True
    ws.Cells(1, LABEL_COLS + 1).HorizontalAlignment = xlCenter
    ws.Cells(1, LABEL_COLS + 1).VerticalAlignment = xlCenter

    ' Column headers - merged rows 2-3 for labels (lighter gray)
    Dim headerColor As Long
    headerColor = RGB(96, 96, 96)

    ws.Range("A2:A3").Merge
    ws.Cells(2, 1).Value = "Project"
    ws.Cells(2, 1).Font.Size = 12
    ws.Cells(2, 1).Font.Bold = True
    ws.Cells(2, 1).HorizontalAlignment = xlCenter
    ws.Cells(2, 1).VerticalAlignment = xlCenter
    ws.Cells(2, 1).WrapText = True
    ws.Cells(2, 1).Interior.Color = headerColor
    ws.Cells(2, 1).Font.Color = RGB(255, 255, 255)

    ws.Range("B2:B3").Merge
    ws.Cells(2, 2).Value = "Task"
    ws.Cells(2, 2).Font.Size = 12
    ws.Cells(2, 2).Font.Bold = True
    ws.Cells(2, 2).HorizontalAlignment = xlCenter
    ws.Cells(2, 2).VerticalAlignment = xlCenter
    ws.Cells(2, 2).WrapText = True
    ws.Cells(2, 2).Interior.Color = headerColor
    ws.Cells(2, 2).Font.Color = RGB(255, 255, 255)

    ws.Range("C2:C3").Merge
    ws.Cells(2, 3).Value = "%"
    ws.Cells(2, 3).Font.Size = 12
    ws.Cells(2, 3).Font.Bold = True
    ws.Cells(2, 3).HorizontalAlignment = xlCenter
    ws.Cells(2, 3).VerticalAlignment = xlCenter
    ws.Cells(2, 3).Interior.Color = headerColor
    ws.Cells(2, 3).Font.Color = RGB(255, 255, 255)

    ws.Range("D2:D3").Merge
    ws.Cells(2, 4).Value = "Owner"
    ws.Cells(2, 4).Font.Size = 12
    ws.Cells(2, 4).Font.Bold = True
    ws.Cells(2, 4).HorizontalAlignment = xlCenter
    ws.Cells(2, 4).VerticalAlignment = xlCenter
    ws.Cells(2, 4).WrapText = True
    ws.Cells(2, 4).Interior.Color = headerColor
    ws.Cells(2, 4).Font.Color = RGB(255, 255, 255)

    ' Week headers
    prevMonth = ""
    monthStartCol = LABEL_COLS + 1

    For i = 0 To totalWeeks - 1
        weekStart = DateAdd("d", i * 7, startDate)
        ' Friday of this week
        fridayDate = weekStart + (5 - Weekday(weekStart, vbMonday))
        If fridayDate < weekStart Then fridayDate = fridayDate + 7

        currentMonth = Format(weekStart, "mmmm, yyyy")

        ' DD/MM of Friday - consistent format
        ws.Cells(3, LABEL_COLS + 1 + i).Value = Format(fridayDate, "d-mmm")
        ws.Cells(3, LABEL_COLS + 1 + i).HorizontalAlignment = xlCenter
        ws.Cells(3, LABEL_COLS + 1 + i).Font.Size = 9
        ws.Cells(3, LABEL_COLS + 1 + i).Interior.Color = RGB(232, 240, 254)

        ' Month tracking
        If currentMonth <> prevMonth Then
            ' Record month boundary for vertical lines
            If prevMonth <> "" Then
                monthBoundaries(boundaryCount) = LABEL_COLS + 1 + i
                boundaryCount = boundaryCount + 1
            End If

            If prevMonth <> "" And LABEL_COLS + i > monthStartCol Then
                ws.Range(ws.Cells(2, monthStartCol), ws.Cells(2, LABEL_COLS + i)).Merge
                ws.Cells(2, monthStartCol).Value = prevMonth
                ws.Cells(2, monthStartCol).HorizontalAlignment = xlCenter
                ws.Cells(2, monthStartCol).Interior.Color = RGB(66, 133, 244)
                ws.Cells(2, monthStartCol).Font.Color = RGB(255, 255, 255)
                ws.Cells(2, monthStartCol).Font.Bold = True
                ws.Cells(2, monthStartCol).Font.Size = 16
            End If
            monthStartCol = LABEL_COLS + 1 + i
            prevMonth = currentMonth
        End If
    Next i

    ' Last month
    If LABEL_COLS + totalWeeks >= monthStartCol Then
        ws.Range(ws.Cells(2, monthStartCol), ws.Cells(2, LABEL_COLS + totalWeeks)).Merge
    End If
    ws.Cells(2, monthStartCol).Value = prevMonth
    ws.Cells(2, monthStartCol).HorizontalAlignment = xlCenter
    ws.Cells(2, monthStartCol).Interior.Color = RGB(66, 133, 244)
    ws.Cells(2, monthStartCol).Font.Color = RGB(255, 255, 255)
    ws.Cells(2, monthStartCol).Font.Bold = True
    ws.Cells(2, monthStartCol).Font.Size = 16

    ' Store month boundaries for later use
    ReDim Preserve monthBoundaries(0 To boundaryCount)

    ' Clean borders for label columns
    ws.Range("A2:A3").Borders.LineStyle = xlContinuous
    ws.Range("A2:A3").Borders.Color = headerColor
    ws.Range("A2:A3").Borders.Weight = xlThin

    ws.Range("B2:B3").Borders.LineStyle = xlContinuous
    ws.Range("B2:B3").Borders.Color = headerColor
    ws.Range("B2:B3").Borders.Weight = xlThin

    ws.Range("C2:C3").Borders.LineStyle = xlContinuous
    ws.Range("C2:C3").Borders.Color = headerColor
    ws.Range("C2:C3").Borders.Weight = xlThin

    ws.Range("D2:D3").Borders.LineStyle = xlContinuous
    ws.Range("D2:D3").Borders.Color = headerColor
    ws.Range("D2:D3").Borders.Weight = xlThin
End Sub

'===============================================================================
Private Function BuildTaskBars(wsData As Worksheet, wsTimeline As Worksheet, _
                                startDate As Date, endDate As Date, _
                                totalWeeks As Integer, lastRow As Long) As Integer
    Dim i As Long, row As Integer
    Dim taskName As String, owner As String, project As String
    Dim taskStart As Variant, taskEnd As Variant, origEnd As Variant
    Dim percentComplete As Double
    Dim barsCreated As Integer
    Dim barColor As Long
    Dim startWeek As Double, endWeek As Double
    Dim clippedStart As Double, clippedEnd As Double
    Dim barLeft As Double, barWidth As Double, barTop As Double
    Dim shp As Shape, progressShp As Shape, termShp As Shape
    Dim colWidth As Double, baseLeft As Double
    Dim origEndWeek As Double, mainBarWidth As Double
    Dim slipLeft As Double, slipWidth As Double, progressWidth As Double
    Dim hasSlip As Boolean

    barsCreated = 0
    baseLeft = wsTimeline.Cells(HEADER_ROWS + 1, LABEL_COLS + 1).Left
    colWidth = wsTimeline.Columns(LABEL_COLS + 1).Width
    If colWidth < 10 Then colWidth = 50

    For i = 2 To lastRow
        row = HEADER_ROWS + (i - 1)
        wsTimeline.Rows(row).RowHeight = TASK_ROW_HEIGHT

        taskName = "" & wsData.Cells(i, COL_TASK_NAME).Value
        project = "" & wsData.Cells(i, COL_PROJECT).Value
        owner = "" & wsData.Cells(i, COL_OWNER).Value
        taskStart = wsData.Cells(i, COL_START_DATE).Value
        taskEnd = wsData.Cells(i, COL_END_DATE).Value
        origEnd = wsData.Cells(i, COL_ORIG_END).Value

        On Error Resume Next
        percentComplete = Val(wsData.Cells(i, COL_PERCENT).Value)
        On Error GoTo 0

        ' Labels - 12pt font
        wsTimeline.Cells(row, 1).Value = project
        wsTimeline.Cells(row, 1).Font.Size = 12
        wsTimeline.Cells(row, 1).VerticalAlignment = xlCenter
        wsTimeline.Cells(row, 1).WrapText = True

        wsTimeline.Cells(row, 2).Value = taskName
        wsTimeline.Cells(row, 2).Font.Bold = True
        wsTimeline.Cells(row, 2).Font.Size = 12
        wsTimeline.Cells(row, 2).VerticalAlignment = xlCenter
        wsTimeline.Cells(row, 2).WrapText = True

        ' Percentage column
        wsTimeline.Cells(row, 3).NumberFormat = "@"  ' Text format
        If percentComplete > 0 Then
            wsTimeline.Cells(row, 3).Value = CInt(percentComplete) & "%"
        End If
        wsTimeline.Cells(row, 3).Font.Size = 12
        wsTimeline.Cells(row, 3).Font.Bold = True
        wsTimeline.Cells(row, 3).HorizontalAlignment = xlCenter
        wsTimeline.Cells(row, 3).VerticalAlignment = xlCenter

        wsTimeline.Cells(row, 4).Value = owner
        wsTimeline.Cells(row, 4).Font.Color = RGB(80, 80, 80)
        wsTimeline.Cells(row, 4).Font.Size = 12
        wsTimeline.Cells(row, 4).WrapText = True
        wsTimeline.Cells(row, 4).VerticalAlignment = xlCenter

        ' Alternate row colors
        If (i Mod 2) = 0 Then
            wsTimeline.Range(wsTimeline.Cells(row, 1), wsTimeline.Cells(row, LABEL_COLS + totalWeeks)).Interior.Color = ROW_COLOR_2
        Else
            wsTimeline.Range(wsTimeline.Cells(row, 1), wsTimeline.Cells(row, LABEL_COLS + totalWeeks)).Interior.Color = ROW_COLOR_1
        End If

        ' Row bottom border
        wsTimeline.Range(wsTimeline.Cells(row, 1), wsTimeline.Cells(row, LABEL_COLS + totalWeeks)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        wsTimeline.Range(wsTimeline.Cells(row, 1), wsTimeline.Cells(row, LABEL_COLS + totalWeeks)).Borders(xlEdgeBottom).Color = RGB(230, 230, 230)

        ' Clean column borders for project info columns
        wsTimeline.Cells(row, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        wsTimeline.Cells(row, 1).Borders(xlEdgeRight).Color = RGB(200, 200, 200)
        wsTimeline.Cells(row, 1).Borders(xlEdgeRight).Weight = xlThin

        wsTimeline.Cells(row, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
        wsTimeline.Cells(row, 2).Borders(xlEdgeRight).Color = RGB(200, 200, 200)
        wsTimeline.Cells(row, 2).Borders(xlEdgeRight).Weight = xlThin

        wsTimeline.Cells(row, 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        wsTimeline.Cells(row, 3).Borders(xlEdgeRight).Color = RGB(200, 200, 200)
        wsTimeline.Cells(row, 3).Borders(xlEdgeRight).Weight = xlThin

        wsTimeline.Cells(row, 4).Borders(xlEdgeRight).LineStyle = xlContinuous
        wsTimeline.Cells(row, 4).Borders(xlEdgeRight).Color = RGB(200, 200, 200)
        wsTimeline.Cells(row, 4).Borders(xlEdgeRight).Weight = xlThin

        ' Project color indicator bar on left edge
        Dim colorBar As Shape
        On Error Resume Next
        Set colorBar = wsTimeline.Shapes.AddShape(msoShapeRectangle, _
            wsTimeline.Cells(row, 1).Left, _
            wsTimeline.Cells(row, 1).Top, _
            4, _
            TASK_ROW_HEIGHT)
        If Not colorBar Is Nothing Then
            colorBar.Fill.ForeColor.RGB = GetProjectColor(project)
            colorBar.Line.Visible = msoFalse
        End If
        On Error GoTo 0

        If IsEmpty(taskStart) Or IsEmpty(taskEnd) Then GoTo NextTask
        If Not IsDate(taskStart) Or Not IsDate(taskEnd) Then GoTo NextTask

        startWeek = (CDate(taskStart) - startDate) / 7
        endWeek = (CDate(taskEnd) - startDate) / 7

        If endWeek <= 0 Or startWeek >= totalWeeks Then GoTo NextTask

        clippedStart = IIf(startWeek < 0, 0, startWeek)
        clippedEnd = IIf(endWeek > totalWeeks, totalWeeks, endWeek)

        If clippedEnd <= clippedStart Then GoTo NextTask

        ' Check slip
        hasSlip = False
        origEndWeek = clippedEnd
        If Not IsEmpty(origEnd) And IsDate(origEnd) Then
            If CDate(taskEnd) > CDate(origEnd) Then
                hasSlip = True
                origEndWeek = (CDate(origEnd) - startDate) / 7
                If origEndWeek < clippedStart Then origEndWeek = clippedStart
                If origEndWeek > clippedEnd Then origEndWeek = clippedEnd
            End If
        End If

        barLeft = baseLeft + (clippedStart * colWidth)
        barTop = wsTimeline.Cells(row, 1).Top + BAR_TOP_MARGIN
        barColor = GetProjectColor(project)

        ' Main bar width (excluding slip)
        If hasSlip Then
            mainBarWidth = (origEndWeek - clippedStart) * colWidth
        Else
            mainBarWidth = (clippedEnd - clippedStart) * colWidth
        End If
        If mainBarWidth < 15 Then mainBarWidth = 15

        ' Create main bar
        On Error Resume Next
        Set shp = wsTimeline.Shapes.AddShape(msoShapeRoundedRectangle, barLeft, barTop, mainBarWidth, BAR_HEIGHT)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextTask
        End If
        On Error GoTo 0

        shp.Fill.ForeColor.RGB = barColor
        shp.Line.ForeColor.RGB = DarkenColor(barColor, 0.2)
        shp.Line.Weight = 1

        ' Progress indicator (darker portion)
        If percentComplete > 0 And percentComplete <= 100 Then
            progressWidth = mainBarWidth * (percentComplete / 100)
            If progressWidth >= 5 Then
                On Error Resume Next
                Set progressShp = wsTimeline.Shapes.AddShape(msoShapeRoundedRectangle, barLeft, barTop, progressWidth, BAR_HEIGHT)
                If Err.Number = 0 And Not progressShp Is Nothing Then
                    progressShp.Fill.ForeColor.RGB = DarkenColor(barColor, 0.15)
                    progressShp.Line.Visible = msoFalse
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If

        ' Slip indicator - dashed horizontal line with termination marker
        If hasSlip And clippedEnd > origEndWeek Then
            slipLeft = baseLeft + (origEndWeek * colWidth)
            slipWidth = (clippedEnd - origEndWeek) * colWidth

            If slipWidth > 3 Then
                On Error Resume Next
                ' Dashed horizontal line (centered vertically in bar area)
                Dim lineTop As Double
                Dim slipLine As Shape
                lineTop = barTop + (BAR_HEIGHT / 2)
                Set slipLine = wsTimeline.Shapes.AddLine(slipLeft, lineTop, slipLeft + slipWidth, lineTop)
                If Err.Number = 0 And Not slipLine Is Nothing Then
                    slipLine.Line.ForeColor.RGB = barColor
                    slipLine.Line.Weight = 2
                    slipLine.Line.DashStyle = msoLineDash
                End If

                ' Termination marker (vertical bar at the end)
                Set termShp = wsTimeline.Shapes.AddShape(msoShapeRectangle, slipLeft + slipWidth - 2, barTop + 4, 3, BAR_HEIGHT - 8)
                If Err.Number = 0 And Not termShp Is Nothing Then
                    termShp.Fill.ForeColor.RGB = barColor
                    termShp.Line.Visible = msoFalse
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If

        barsCreated = barsCreated + 1

NextTask:
    Next i

    BuildTaskBars = barsCreated
End Function

'===============================================================================
Private Sub AddMonthDividers(ws As Worksheet, startDate As Date, totalWeeks As Integer, taskCount As Long)
    Dim i As Integer
    Dim weekStart As Date
    Dim currentMonth As String
    Dim prevMonth As String
    Dim colLeft As Double
    Dim lineTop As Double
    Dim lineHeight As Double
    Dim shp As Shape

    On Error Resume Next

    prevMonth = ""
    lineTop = ws.Cells(HEADER_ROWS + 1, 1).Top
    lineHeight = taskCount * TASK_ROW_HEIGHT

    For i = 0 To totalWeeks - 1
        weekStart = DateAdd("d", i * 7, startDate)
        currentMonth = Format(weekStart, "mmmm yyyy")

        If currentMonth <> prevMonth And prevMonth <> "" Then
            ' Draw vertical line at month boundary
            colLeft = ws.Cells(HEADER_ROWS + 1, LABEL_COLS + 1 + i).Left
            Set shp = ws.Shapes.AddShape(msoShapeRectangle, colLeft, lineTop, 1, lineHeight)
            If Not shp Is Nothing Then
                shp.Fill.ForeColor.RGB = RGB(66, 133, 244)
                shp.Line.Visible = msoFalse
            End If
        End If
        prevMonth = currentMonth
    Next i

    On Error GoTo 0
End Sub

'===============================================================================
Private Sub AddTodayMarker(ws As Worksheet, startDate As Date, totalWeeks As Integer, taskCount As Long)
    Dim todayWeek As Double
    Dim markerLeft As Double, markerTop As Double, markerHeight As Double
    Dim shp As Shape
    Dim colWidth As Double, baseLeft As Double

    On Error Resume Next

    todayWeek = (Date - startDate) / 7

    If todayWeek >= 0 And todayWeek < totalWeeks Then
        baseLeft = ws.Cells(HEADER_ROWS + 1, LABEL_COLS + 1).Left
        colWidth = ws.Columns(LABEL_COLS + 1).Width
        If colWidth < 10 Then colWidth = 50

        markerLeft = baseLeft + (todayWeek * colWidth)
        markerTop = ws.Cells(HEADER_ROWS + 1, 1).Top
        markerHeight = taskCount * TASK_ROW_HEIGHT

        If markerHeight > 0 And markerLeft > 0 Then
            Set shp = ws.Shapes.AddShape(msoShapeRectangle, markerLeft, markerTop, 2, markerHeight)
            If Not shp Is Nothing Then
                shp.Fill.ForeColor.RGB = RGB(255, 87, 34)
                shp.Line.Visible = msoFalse
            End If
        End If

        ' Today label
        Dim weekCol As Integer
        weekCol = LABEL_COLS + 1 + Int(todayWeek)
        If weekCol > 0 Then
            ws.Cells(3, weekCol).Font.Color = RGB(255, 87, 34)
            ws.Cells(3, weekCol).Font.Bold = True
        End If
    End If

    On Error GoTo 0
End Sub

'===============================================================================
Private Sub AddChartBorder(ws As Worksheet, taskCount As Long, totalWeeks As Integer)
    Dim lastRow As Integer, lastCol As Integer
    Dim chartRange As Range

    lastRow = HEADER_ROWS + taskCount
    lastCol = LABEL_COLS + totalWeeks

    ' Empty border row/column
    ws.Rows(lastRow + 1).RowHeight = 6
    ws.Columns(lastCol + 1).ColumnWidth = 1
    ws.Rows(lastRow + 1).Interior.Color = RGB(255, 255, 255)
    ws.Columns(lastCol + 1).Interior.Color = RGB(255, 255, 255)

    ' Outer border
    Set chartRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow + 1, lastCol + 1))
    With chartRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = RGB(180, 180, 180)
        .Weight = xlThin
    End With
    With chartRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = RGB(180, 180, 180)
        .Weight = xlThin
    End With
    With chartRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = RGB(180, 180, 180)
        .Weight = xlThin
    End With
    With chartRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(180, 180, 180)
        .Weight = xlThin
    End With
End Sub

'===============================================================================
Private Sub InitializeColors()
    ' Reset project color tracking
    Set projectColors = New Collection
    colorIndex = 0

    ' Define color palette (10 distinct colors)
    colorPalette(0) = RGB(66, 133, 244)   ' Blue
    colorPalette(1) = RGB(251, 188, 5)    ' Yellow/Gold
    colorPalette(2) = RGB(234, 67, 53)    ' Red
    colorPalette(3) = RGB(156, 39, 176)   ' Purple
    colorPalette(4) = RGB(52, 168, 83)    ' Green
    colorPalette(5) = RGB(255, 112, 67)   ' Orange
    colorPalette(6) = RGB(0, 172, 193)    ' Cyan
    colorPalette(7) = RGB(171, 71, 188)   ' Light Purple
    colorPalette(8) = RGB(124, 179, 66)   ' Lime
    colorPalette(9) = RGB(255, 167, 38)   ' Amber
End Sub

'===============================================================================
Private Function GetProjectColor(project As String) As Long
    Dim p As String
    Dim existingColor As Variant

    p = LCase(Trim(project))
    If p = "" Then
        GetProjectColor = RGB(74, 101, 114)  ' Default gray for empty
        Exit Function
    End If

    ' Check if project already has a color assigned
    On Error Resume Next
    existingColor = projectColors(p)
    On Error GoTo 0

    If Not IsEmpty(existingColor) Then
        GetProjectColor = existingColor
    Else
        ' Assign next color from palette
        GetProjectColor = colorPalette(colorIndex Mod 10)
        projectColors.Add GetProjectColor, p
        colorIndex = colorIndex + 1
    End If
End Function

'===============================================================================
Private Function DarkenColor(baseColor As Long, amount As Double) As Long
    Dim r As Integer, g As Integer, b As Integer
    r = baseColor Mod 256
    g = (baseColor \ 256) Mod 256
    b = (baseColor \ 65536) Mod 256
    DarkenColor = RGB(Int(r * (1 - amount)), Int(g * (1 - amount)), Int(b * (1 - amount)))
End Function

'===============================================================================
Private Function LightenColor(baseColor As Long, amount As Double) As Long
    Dim r As Integer, g As Integer, b As Integer
    r = baseColor Mod 256
    g = (baseColor \ 256) Mod 256
    b = (baseColor \ 65536) Mod 256
    LightenColor = RGB(Int(r + (255 - r) * amount), Int(g + (255 - g) * amount), Int(b + (255 - b) * amount))
End Function
