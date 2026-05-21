Attribute VB_Name = "TicketDashboardModule"
Option Explicit

' ============================================================================
' TICKET ANALYST DASHBOARD - COMPLETE VBA MODULE
' ============================================================================
' Data Sheet Columns Required:
' Ticket ID, Date, Status, Priority, Category, Assigned Agent, 
' Resolution Time (Hours), Source, SLA Status, Satisfaction Rating
' ============================================================================

Sub BuildTicketDashboard()

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim wsDash As Worksheet
    Dim lo As ListObject
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRng As Range

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("Data")

    ' STEP 1: Convert Data sheet range to Excel Table named tblTickets
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Set dataRng = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))

    On Error Resume Next
    wsData.ListObjects("tblTickets").Delete
    On Error GoTo 0

    Set lo = wsData.ListObjects.Add(xlSrcRange, dataRng, , xlYes)
    lo.Name = "tblTickets"
    lo.TableStyle = "TableStyleMedium2"

    ' STEP 2: Delete and recreate Pivot sheet
    On Error Resume Next
    wb.Sheets("Pivot").Delete
    On Error GoTo 0

    Set wsPivot = wb.Sheets.Add(After:=wsData)
    wsPivot.Name = "Pivot"


    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.ListObjects("tblTickets").Range)

    Call CreatePivot_StatusTrend(pc, wsPivot)
    Call CreatePivot_PriorityBreakdown(pc, wsPivot)
    Call CreatePivot_CategoryAnalysis(pc, wsPivot)
    Call CreatePivot_AgentPerformance(pc, wsPivot)
    Call CreatePivot_ResolutionTime(pc, wsPivot)
    Call CreatePivot_SourceChannel(pc, wsPivot)
    Call CreatePivot_SLACompliance(pc, wsPivot)
    Call CreatePivot_CustomerSatisfaction(pc, wsPivot)
    Call CreatePivot_KPI(pc, wsPivot)

    ' STEP 3: Delete and recreate Dashboard sheet
    On Error Resume Next
    wb.Sheets("Dashboard").Delete
    On Error GoTo 0

    Set wsDash = wb.Sheets.Add(Before:=wsData)
    wsDash.Name = "Dashboard"

    wsDash.Activate
    ActiveWindow.DisplayGridlines = False

    ' Background and font
    wsDash.Cells.Interior.Color = RGB(240, 242, 245)
    wsDash.Cells.Font.Name = "Calibri"

    ' Column widths: A-T = 8.2, U = 0.6, V-X = 8.2
    Dim i As Integer
    For i = 1 To 20
        wsDash.Columns(i).ColumnWidth = 8.2
    Next i
    wsDash.Columns(21).ColumnWidth = 0.6
    For i = 22 To 24
        wsDash.Columns(i).ColumnWidth = 8.2
    Next i
    wsDash.Rows.RowHeight = 13.5
    wsDash.Rows(11).RowHeight = 0

    ' STEP 3b: Header Banner
    Dim hdrLeft   As Single: hdrLeft   = wsDash.Range("A1").Left
    Dim hdrTop    As Single: hdrTop    = wsDash.Range("A1").Top
    Dim hdrWidth  As Single: hdrWidth  = wsDash.Range("A1:X1").Width
    Dim hdrHeight As Single: hdrHeight = wsDash.Range("A1:A3").Height


    Dim hdrShape As Shape
    Set hdrShape = wsDash.Shapes.AddShape(msoShapeRectangle, _
        hdrLeft, hdrTop, hdrWidth, hdrHeight)
    With hdrShape
        .Fill.ForeColor.RGB = RGB(21, 67, 96)
        .Fill.Transparency  = 0
        .Line.Visible       = msoFalse
        .Name               = "HeaderBanner"
    End With
    With hdrShape.TextFrame
        .MarginLeft   = 0
        .MarginRight  = 0
        .MarginTop    = 0
        .MarginBottom = 0
        .Characters.Text           = "Ticket Support Analytics Dashboard"
        .Characters.Font.Bold      = True
        .Characters.Font.Color     = RGB(255, 255, 255)
        .Characters.Font.Size      = 36
        .Characters.Font.Name      = "Calibri"
        .HorizontalAlignment       = xlHAlignCenter
        .VerticalAlignment         = xlVAlignCenter
    End With

    ' STEP 4: KPI Cards
    Call CreateKPICards(wsDash, wsPivot)

    ' STEP 5: Charts
    Call CreateCharts(wsDash, wsPivot)

    ' STEP 6: Slicers
    Call CreateSlicers(wb, wsDash, wsPivot)

    wsDash.Activate
    wsDash.Range("A1").Select

    ' STEP 7: Turn screen on, then apply zoom with force-render cycle
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    ActiveWindow.Zoom = 100
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True

    ActiveWindow.Zoom = 70
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True

    wb.RefreshAll

    Call InjectDashboardEvent(wb)
    Call UpdateKPIValues(wsDash, wsPivot)


    Call FixDonutCenter

    MsgBox "Ticket Analytics Dashboard built successfully!", vbInformation, "Dashboard Ready"

End Sub

' ============================================================================
' INJECT DASHBOARD EVENT HANDLERS
' ============================================================================
Private Sub InjectDashboardEvent(wb As Workbook)
    On Error Resume Next

    Dim cm As Object
    Dim vbc As Object
    For Each vbc In wb.VBProject.VBComponents
        If vbc.Properties("Name") = "Dashboard" Then
            Set cm = vbc.CodeModule
            Exit For
        End If
    Next vbc

    If cm Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If

    Dim i As Long
    Dim startLine As Long, endLine As Long
    For i = cm.CountOfLines To 1 Step -1
        Dim lineText As String
        lineText = cm.Lines(i, 1)
        If InStr(lineText, "Worksheet_Calculate") > 0 Or _
           InStr(lineText, "Worksheet_PivotTableUpdate") > 0 Then
            startLine = i
            endLine = i
            Do While endLine < cm.CountOfLines
                endLine = endLine + 1
                If InStr(cm.Lines(endLine, 1), "End Sub") > 0 Then Exit Do
            Loop
            cm.DeleteLines startLine, endLine - startLine + 1
        End If
    Next i

    Dim calcEvent As String
    calcEvent = "Private Sub Worksheet_Calculate()" & vbCrLf & _
                "    On Error Resume Next" & vbCrLf & _
                "    Dim wsPivot As Worksheet" & vbCrLf & _
                "    Set wsPivot = ThisWorkbook.Sheets(""Pivot"")" & vbCrLf & _
                "    If Not wsPivot Is Nothing Then" & vbCrLf & _
                "        Call TicketDashboardModule.UpdateKPIValues(Me, wsPivot)" & vbCrLf & _
                "    End If" & vbCrLf & _
                "    On Error GoTo 0" & vbCrLf & _
                "End Sub"


    Dim pvtEvent As String
    pvtEvent = "Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)" & vbCrLf & _
               "    On Error Resume Next" & vbCrLf & _
               "    Dim wsPivot As Worksheet" & vbCrLf & _
               "    Set wsPivot = ThisWorkbook.Sheets(""Pivot"")" & vbCrLf & _
               "    If Not wsPivot Is Nothing Then" & vbCrLf & _
               "        Call TicketDashboardModule.UpdateKPIValues(Me, wsPivot)" & vbCrLf & _
               "    End If" & vbCrLf & _
               "    On Error GoTo 0" & vbCrLf & _
               "End Sub"

    Dim insertAt As Long
    insertAt = cm.CountOfLines + 1
    cm.InsertLines insertAt, calcEvent
    cm.InsertLines cm.CountOfLines + 1, pvtEvent

    On Error GoTo 0
End Sub

' ============================================================================
' PIVOT TABLE CREATION SUBS
' ============================================================================
Private Sub CreatePivot_StatusTrend(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A2"), _
        TableName:="pvtStatusTrend")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Date").Orientation = xlRowField
        .PivotFields("Date").Position = 1
        
        On Error Resume Next
        .PivotFields("Date").LabelRange.Group Start:=True, End:=True, _
            Periods:=Array(False, False, False, False, True, False, False)
        On Error GoTo 0
        
        .PivotFields("Status").Orientation = xlColumnField
        .PivotFields("Status").Position = 1

        With .PivotFields("Ticket ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count of Tickets"
            .NumberFormat = "#,##0"
        End With


        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_PriorityBreakdown(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A25"), _
        TableName:="pvtPriority")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Priority").Orientation = xlRowField
        .PivotFields("Priority").Position = 1

        With .PivotFields("Ticket ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Tickets by Priority"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_CategoryAnalysis(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("F2"), _
        TableName:="pvtCategory")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Category").Orientation = xlRowField
        .PivotFields("Category").Position = 1

        With .PivotFields("Ticket ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Tickets by Category"
            .NumberFormat = "#,##0"
        End With


        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_AgentPerformance(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("F25"), _
        TableName:="pvtAgent")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Assigned Agent").Orientation = xlRowField
        .PivotFields("Assigned Agent").Position = 1

        With .PivotFields("Ticket ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Tickets per Agent"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_ResolutionTime(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("K2"), _
        TableName:="pvtResolution")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Status").Orientation = xlRowField
        .PivotFields("Status").Position = 1

        With .PivotFields("Resolution Time (Hours)")
            .Orientation = xlDataField
            .Function = xlAverage
            .Name = "Avg Resolution Time"
            .NumberFormat = "0.0"
        End With


        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_SourceChannel(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("K20"), _
        TableName:="pvtSource")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Source").Orientation = xlRowField
        .PivotFields("Source").Position = 1

        With .PivotFields("Ticket ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Tickets by Source"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_SLACompliance(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("P2"), _
        TableName:="pvtSLA")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("SLA Status").Orientation = xlRowField
        .PivotFields("SLA Status").Position = 1

        With .PivotFields("Ticket ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "SLA Compliance"
            .NumberFormat = "#,##0"
        End With


        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_CustomerSatisfaction(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("P20"), _
        TableName:="pvtCSAT")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Satisfaction Rating").Orientation = xlRowField
        .PivotFields("Satisfaction Rating").Position = 1

        With .PivotFields("Ticket ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Customer Satisfaction"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_KPI(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("U5"), _
        TableName:="PT_KPI")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        With .PivotFields("Ticket ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Total Tickets"
            .NumberFormat = "#,##0"
        End With


        With .PivotFields("Resolution Time (Hours)")
            .Orientation = xlDataField
            .Function = xlAverage
            .Name = "Avg Resolution"
            .NumberFormat = "0.0"
        End With

        With .PivotFields("Satisfaction Rating")
            .Orientation = xlDataField
            .Function = xlAverage
            .Name = "Avg CSAT"
            .NumberFormat = "0.0"
        End With

        On Error Resume Next
        .DataPivotField.Orientation = xlRowField
        .DataPivotField.Position = 1
        On Error GoTo 0

        .RowAxisLayout xlTabularRow
        .DisplayFieldCaptions = False
        .ShowTableStyleRowStripes = False
        .ColumnGrand = False
        .RowGrand = False
    End With
End Sub

' ============================================================================
' KPI CARDS CREATION
' ============================================================================
Private Sub CreateKPICards(ws As Worksheet, wsPivot As Worksheet)

    Const KPI_ROW As Long = 200
    ws.Cells(KPI_ROW, 1).Formula = "=Pivot!V5"
    ws.Cells(KPI_ROW, 2).Formula = "=Pivot!V6"
    ws.Cells(KPI_ROW, 3).Formula = "=Pivot!V7"
    ws.Rows(KPI_ROW).Hidden = True

    Dim kpiTitles(5) As String
    kpiTitles(0) = "Total Tickets"
    kpiTitles(1) = "Open Tickets"
    kpiTitles(2) = "Resolved Tickets"
    kpiTitles(3) = "Avg Resolution (Hrs)"
    kpiTitles(4) = "Avg CSAT Score"
    kpiTitles(5) = "SLA Compliance %"

    Dim kpiCells(5) As String
    kpiCells(0) = "=Pivot!V5"
    kpiCells(1) = "=COUNTIFS(Data[Status],""Open"")"
    kpiCells(2) = "=COUNTIFS(Data[Status],""Resolved"")+COUNTIFS(Data[Status],""Closed"")"
    kpiCells(3) = "=Pivot!V6"
    kpiCells(4) = "=Pivot!V7"
    kpiCells(5) = "=IF(Pivot!V5>0,COUNTIFS(Data[SLA Status],""Met"")/Pivot!V5,0)"


    Dim accentColors(5) As Long
    accentColors(0) = RGB(37, 99, 235)
    accentColors(1) = RGB(245, 158, 11)
    accentColors(2) = RGB(5, 150, 105)
    accentColors(3) = RGB(124, 58, 237)
    accentColors(4) = RGB(220, 38, 38)
    accentColors(5) = RGB(21, 67, 96)

    Dim areaLeft  As Single: areaLeft  = ws.Range("A1").Left + 5
    Dim areaWidth As Single: areaWidth = ws.Range("A1:T1").Width - 10
    Dim cardGap   As Single: cardGap   = 8
    Dim cardW     As Single: cardW     = Int((areaWidth - 5 * cardGap) / 6)
    Dim cardH     As Single: cardH     = 85
    Dim accentH   As Single: accentH   = 12
    Dim cardTop   As Single: cardTop   = ws.Range("A4").Top + 4

    Dim k As Integer
    For k = 0 To 5
        Dim cLeft As Single
        cLeft = areaLeft + k * (cardW + cardGap)

        Dim cardShape As Shape
        Set cardShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, cLeft, cardTop, cardW, cardH)
        With cardShape
            .Name                  = "KPI_Card_" & k
            .Fill.ForeColor.RGB    = RGB(255, 255, 255)
            .Fill.Transparency     = 0
            .Line.ForeColor.RGB    = RGB(218, 224, 235)
            .Line.Weight           = 0.75
            .Shadow.Type           = msoShadow21
            .Shadow.Transparency   = 0.75
            .Shadow.OffsetX        = 1
            .Shadow.OffsetY        = 2
            .Shadow.Size           = 100
            .Shadow.Blur           = 4
        End With

        Dim acBar As Shape
        Set acBar = ws.Shapes.AddShape(msoShapeRectangle, cLeft, cardTop, cardW, accentH)
        With acBar
            .Name               = "KPI_Accent_" & k
            .Fill.ForeColor.RGB = accentColors(k)
            .Fill.Transparency  = 0
            .Line.Visible       = msoFalse
        End With


        Dim titleBox As Shape
        Set titleBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            cLeft + 6, cardTop + accentH + 5, cardW - 12, 22)
        With titleBox
            .Name          = "KPI_Title_" & k
            .Line.Visible  = msoFalse
            .Fill.Visible  = msoFalse
            With .TextFrame
                .Characters.Text           = kpiTitles(k)
                .Characters.Font.Name      = "Calibri"
                .Characters.Font.Size      = 16
                .Characters.Font.Bold      = True
                .Characters.Font.Color     = RGB(10, 10, 10)
                .HorizontalAlignment       = xlHAlignCenter
                .VerticalAlignment         = xlVAlignCenter
            End With
        End With

        Dim valTop  As Single: valTop  = cardTop + accentH + 30
        Dim valH    As Single: valH    = cardH - accentH - 33
        Dim valBox  As Shape
        Set valBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            cLeft + 4, valTop, cardW - 8, valH)
        With valBox
            .Name          = "KPI_Value_" & k
            .Line.Visible  = msoFalse
            .Fill.Visible  = msoFalse
            .DrawingObject.Formula = kpiCells(k)
            With .TextFrame
                .Characters.Font.Name      = "Calibri"
                .Characters.Font.Size      = 22
                .Characters.Font.Bold      = True
                .Characters.Font.Color     = RGB(15, 32, 65)
                .HorizontalAlignment       = xlHAlignCenter
                .VerticalAlignment         = xlVAlignCenter
            End With
        End With

    Next k

End Sub

' ============================================================================
' UPDATE KPI VALUES (Stub)
' ============================================================================
Private Sub UpdateKPIValues(ws As Worksheet, wsPivot As Worksheet)
    ' Nothing to do - shapes are formula-linked to Pivot cells
End Sub


' ============================================================================
' CHARTS CREATION
' ============================================================================
Private Sub CreateCharts(ws As Worksheet, wsPivot As Worksheet)

    Const CHART_COLOR_BG   As Long = 16777215
    Const CHART_COLOR_NAVY As Long = 984097

    Dim startL   As Single: startL   = ws.Range("A1").Left + 5
    Dim contentW As Single: contentW = ws.Range("A1:T1").Width - 10
    Dim cGap     As Single: cGap     = 8

    Dim rw1Top  As Single: rw1Top  = ws.Range("A12").Top + 3
    Dim rw1H    As Single: rw1H    = ws.Range("A12:A26").Height - 6
    Dim rw2Top  As Single: rw2Top  = ws.Range("A27").Top + 3
    Dim rw2H    As Single: rw2H    = ws.Range("A27:A42").Height - 6

    Dim cht1W As Single: cht1W = Int(contentW * 0.48)
    Dim cht2W As Single: cht2W = contentW - cht1W - cGap
    Dim cht3W As Single: cht3W = Int(contentW * 0.28)
    Dim cht4W As Single: cht4W = Int(contentW * 0.35)
    Dim cht5W As Single: cht5W = contentW - cht3W - cht4W - 2 * cGap

    Dim c1L As Single: c1L = startL
    Dim c2L As Single: c2L = c1L + cht1W + cGap
    Dim c3L As Single: c3L = startL
    Dim c4L As Single: c4L = c3L + cht3W + cGap
    Dim c5L As Single: c5L = c4L + cht4W + cGap

    ' Chart 1: Ticket Status Trend (Stacked Area Chart)
    Dim cht1 As ChartObject
    Set cht1 = ws.ChartObjects.Add(c1L, rw1Top, cht1W, rw1H)
    cht1.Name = "chtStatusTrend"
    With cht1.Chart
        .ChartType = xlAreaStacked
        .HasTitle  = True
        .ChartTitle.Text           = "Ticket Status Trend Over Time"
        .ChartTitle.Font.Size      = 11
        .ChartTitle.Font.Bold      = True
        .ChartTitle.Font.Name      = "Calibri"
        .ChartTitle.Font.Color     = CHART_COLOR_NAVY
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range("A2").PivotTable.TableRange1
        On Error GoTo 0
        .HasLegend            = True
        .Legend.Position      = xlLegendPositionBottom
        .Legend.Font.Size     = 8
        .Legend.Font.Name     = "Calibri"
        .PlotArea.Interior.Color    = RGB(250, 252, 255)
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color   = CHART_COLOR_BG


        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 234, 242)
        On Error GoTo 0
        On Error Resume Next: .ShowAllFieldButtons = False: On Error GoTo 0
    End With

    ' Chart 2: Priority Breakdown (Donut Chart)
    On Error Resume Next
    ws.ChartObjects("chtPriority").Delete
    On Error GoTo 0

    Dim cht2 As ChartObject
    Set cht2 = ws.ChartObjects.Add(c2L, rw1Top, cht2W, rw1H)
    cht2.Name = "chtPriority"

    With cht2.Chart
        .ChartType = xlDoughnut
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range("A25").PivotTable.TableRange1
        On Error GoTo 0
        .HasTitle                   = True
        .ChartTitle.Text            = "Priority Distribution"
        .ChartTitle.Font.Size       = 11
        .ChartTitle.Font.Bold       = True
        .ChartTitle.Font.Name       = "Calibri"
        .ChartTitle.Font.Color      = CHART_COLOR_NAVY
        .HasLegend                  = True
        .Legend.Position            = xlLegendPositionRight
        .Legend.Font.Size           = 9
        .Legend.Font.Name           = "Calibri"
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color   = CHART_COLOR_BG
        .PlotArea.Interior.Color    = CHART_COLOR_BG
        On Error Resume Next
        .SeriesCollection(1).DoughnutHoleSize = 40
        Dim dColors(3) As Long
        dColors(0) = RGB(220, 38, 38)
        dColors(1) = RGB(245, 158, 11)
        dColors(2) = RGB(37, 99, 235)
        dColors(3) = RGB(5, 150, 105)
        Dim p As Integer
        For p = 1 To .SeriesCollection(1).Points.Count
            If p <= 4 Then _
                .SeriesCollection(1).Points(p).Format.Fill.ForeColor.RGB = dColors(p - 1)
        Next p
        With .SeriesCollection(1)
            .HasDataLabels               = True
            .DataLabels.ShowCategoryName = True
            .DataLabels.ShowPercentage   = True
            .DataLabels.ShowValue        = False
            .DataLabels.NumberFormat     = "0%"
            .DataLabels.Font.Size        = 8
            .DataLabels.Font.Bold        = True
            .DataLabels.Font.Color       = RGB(255, 255, 255)
        End With
        On Error GoTo 0
        On Error Resume Next: .ShowAllFieldButtons = False: On Error GoTo 0
    End With


    Call CenterDonutPlotArea(cht2)

    ' Chart 3: Tickets by Category (Bar Chart)
    Dim cht3 As ChartObject
    Set cht3 = ws.ChartObjects.Add(c3L, rw2Top, cht3W, rw2H)
    cht3.Name = "chtCategory"
    With cht3.Chart
        .ChartType = xlBarClustered
        .HasTitle  = True
        .ChartTitle.Text       = "Tickets by Category"
        .ChartTitle.Font.Size  = 11
        .ChartTitle.Font.Bold  = True
        .ChartTitle.Font.Name  = "Calibri"
        .ChartTitle.Font.Color = CHART_COLOR_NAVY
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range("F2").PivotTable.TableRange1
        On Error GoTo 0
        .HasLegend = False
        .PlotArea.Interior.Color    = RGB(250, 252, 255)
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color   = CHART_COLOR_BG
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 234, 242)
        On Error GoTo 0
        On Error Resume Next
        With .SeriesCollection(1)
            .Format.Fill.ForeColor.RGB = RGB(37, 99, 235)
            .GapWidth                  = 40
            .HasDataLabels             = True
            .DataLabels.NumberFormat   = "#,##0"
            .DataLabels.Font.Size      = 8
            .DataLabels.Font.Bold      = True
            .DataLabels.Font.Color     = RGB(15, 32, 65)
            .DataLabels.Position       = xlLabelPositionOutsideEnd
        End With
        On Error GoTo 0
        On Error Resume Next: .ShowAllFieldButtons = False: On Error GoTo 0
    End With

    ' Chart 4: Agent Performance (Column Chart)
    Dim cht4 As ChartObject
    Set cht4 = ws.ChartObjects.Add(c4L, rw2Top, cht4W, rw2H)
    cht4.Name = "chtAgent"
    With cht4.Chart
        .ChartType = xlColumnClustered
        .HasTitle  = True
        .ChartTitle.Text       = "Agent Performance"
        .ChartTitle.Font.Size  = 11
        .ChartTitle.Font.Bold  = True
        .ChartTitle.Font.Name  = "Calibri"
        .ChartTitle.Font.Color = CHART_COLOR_NAVY
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range("F25").PivotTable.TableRange1
        On Error GoTo 0


        .HasLegend = False
        .PlotArea.Interior.Color    = RGB(250, 252, 255)
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color   = CHART_COLOR_BG
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 234, 242)
        On Error GoTo 0
        On Error Resume Next
        With .SeriesCollection(1)
            .Format.Fill.ForeColor.RGB = RGB(5, 150, 105)
            .GapWidth                  = 55
            .HasDataLabels             = True
            .DataLabels.NumberFormat   = "#,##0"
            .DataLabels.Font.Size      = 8
            .DataLabels.Font.Bold      = True
            .DataLabels.Font.Color     = RGB(15, 32, 65)
            .DataLabels.Position       = xlLabelPositionOutsideEnd
        End With
        On Error GoTo 0
        On Error Resume Next: .ShowAllFieldButtons = False: On Error GoTo 0
    End With

    ' Chart 5: Customer Satisfaction (Column Chart)
    Dim cht5 As ChartObject
    Set cht5 = ws.ChartObjects.Add(c5L, rw2Top, cht5W, rw2H)
    cht5.Name = "chtCSAT"
    With cht5.Chart
        .ChartType = xlColumnClustered
        .HasTitle  = True
        .ChartTitle.Text       = "Customer Satisfaction Rating"
        .ChartTitle.Font.Size  = 11
        .ChartTitle.Font.Bold  = True
        .ChartTitle.Font.Name  = "Calibri"
        .ChartTitle.Font.Color = CHART_COLOR_NAVY
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range("P20").PivotTable.TableRange1
        On Error GoTo 0
        .HasLegend = False
        .PlotArea.Interior.Color    = RGB(250, 252, 255)
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color   = CHART_COLOR_BG
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 234, 242)
        On Error GoTo 0
        On Error Resume Next
        With .SeriesCollection(1)
            .Format.Fill.ForeColor.RGB = RGB(245, 158, 11)
            .GapWidth                  = 50
            .HasDataLabels             = True
            .DataLabels.NumberFormat   = "#,##0"
            .DataLabels.Font.Size      = 8
            .DataLabels.Font.Bold      = True
            .DataLabels.Font.Color     = RGB(180, 110, 0)


            .DataLabels.Position       = xlLabelPositionOutsideEnd
        End With
        On Error GoTo 0
        On Error Resume Next: .ShowAllFieldButtons = False: On Error GoTo 0
    End With

End Sub

' ============================================================================
' SLICERS CREATION
' ============================================================================
Private Sub CreateSlicers(wb As Workbook, ws As Worksheet, wsPivot As Worksheet)

    Dim panelLeft   As Single: panelLeft   = ws.Range("V4").Left
    Dim panelTop    As Single: panelTop    = ws.Range("V4").Top
    Dim panelWidth  As Single: panelWidth  = ws.Range("V4:X4").Width
    Dim panelHeight As Single: panelHeight = ws.Range("V4:V42").Height

    Dim panelBG As Shape
    Set panelBG = ws.Shapes.AddShape(msoShapeRectangle, _
        panelLeft, panelTop, panelWidth, panelHeight)
    With panelBG
        .Name               = "SlicerPanel_BG"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Transparency  = 0
        .Line.ForeColor.RGB = RGB(210, 218, 232)
        .Line.Weight        = 0.75
        .ZOrder msoSendToBack
    End With

    Dim slicerW   As Single: slicerW   = panelWidth - 8
    Dim slicerGap As Single: slicerGap = 8

    Dim slicerH(3) As Single
    slicerH(0) = 120
    slicerH(1) = 90
    slicerH(2) = 100
    slicerH(3) = 120

    Dim slicerCols(3) As Integer
    slicerCols(0) = 1
    slicerCols(1) = 2
    slicerCols(2) = 2
    slicerCols(3) = 1

    Dim slicerLeft     As Single: slicerLeft     = panelLeft + 4
    Dim slicerTopStart As Single: slicerTopStart = panelTop  + 4

    Dim slicerFields(3) As String
    slicerFields(0) = "Status"
    slicerFields(1) = "Priority"
    slicerFields(2) = "Category"
    slicerFields(3) = "Assigned Agent"


    Dim slicerNames(3) As String
    slicerNames(0) = "slcStatus"
    slicerNames(1) = "slcPriority"
    slicerNames(2) = "slcCategory"
    slicerNames(3) = "slcAgent"

    Dim ptNames(7) As String
    ptNames(0) = "pvtStatusTrend"
    ptNames(1) = "pvtPriority"
    ptNames(2) = "pvtCategory"
    ptNames(3) = "pvtAgent"
    ptNames(4) = "pvtResolution"
    ptNames(5) = "pvtSource"
    ptNames(6) = "pvtSLA"
    ptNames(7) = "pvtCSAT"

    Dim existSC As SlicerCache
    Dim scToDelete() As String
    Dim scCount As Integer
    scCount = 0

    For Each existSC In wb.SlicerCaches
        ReDim Preserve scToDelete(scCount)
        scToDelete(scCount) = existSC.Name
        scCount = scCount + 1
    Next existSC

    Dim d As Integer
    For d = 0 To scCount - 1
        On Error Resume Next
        wb.SlicerCaches(scToDelete(d)).Delete
        On Error GoTo 0
    Next d

    Dim f As Integer
    Dim currentTop As Single
    currentTop = slicerTopStart

    For f = 0 To 3
        Dim sc As SlicerCache
        Dim sl As Slicer
        Dim firstPT As PivotTable

        On Error Resume Next
        Set firstPT = wsPivot.PivotTables("pvtStatusTrend")

        Set sc = wb.SlicerCaches.Add2(firstPT, slicerFields(f), slicerNames(f))

        Dim pn As Integer
        For pn = 0 To 7
            On Error Resume Next
            sc.PivotTables.AddPivotTable wsPivot.PivotTables(ptNames(pn))
            On Error GoTo 0
        Next pn


        Set sl = sc.Slicers.Add(ws, , slicerNames(f) & "_visual", _
            slicerFields(f), currentTop, slicerLeft, slicerW, slicerH(f))

        With sl
            .Style           = "SlicerStyleDark4"
            .Left            = slicerLeft
            .Top             = currentTop
            .Width           = slicerW
            .Height          = slicerH(f)
            .NumberOfColumns = slicerCols(f)
            .RowHeight       = 20
        End With

        On Error GoTo 0

        currentTop = currentTop + slicerH(f) + slicerGap
    Next f

End Sub

' ============================================================================
' DONUT CHART CENTER HELPERS
' ============================================================================
Private Sub CenterDonutPlotArea(chtObj As ChartObject)
    On Error Resume Next
    Dim cht  As Chart:  Set cht  = chtObj.Chart
    Dim caW  As Single: caW  = cht.ChartArea.Width
    Dim caH  As Single: caH  = cht.ChartArea.Height

    With cht.PlotArea
        .Left   = 0
        .Top    = 0
        .Width  = caW
        .Height = caH
    End With
    With cht.PlotArea
        .Left   = 0
        .Top    = 0
        .Width  = caW
        .Height = caH
    End With
    On Error GoTo 0
End Sub

Private Sub FixDonutCenter()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    If ws Is Nothing Then Exit Sub
    Call CenterDonutPlotArea(ws.ChartObjects("chtPriority"))
    On Error GoTo 0
End Sub
