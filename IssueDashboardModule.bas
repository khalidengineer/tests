Attribute VB_Name = "IssueDashboardModule"
Option Explicit

' ============================================================================
' ISSUE TRACKING ANALYTICS DASHBOARD - COMPLETE VBA MODULE
' ============================================================================
' Data Sheet Columns Required (24 columns):
' Issue Type | Severity | Current Status | Issue Location | Issue Description |
' Resolver | Resolver Identifier | Reporter | Reporter Identifier |
' Reporter Designation | Reporter Department | Reporter Division |
' Reporter Sub Division | Reporter Location | Resolver Designation |
' Resolver Department | Resolver Division | Resolver Sub Division |
' Resolver Location | Resolved At | Resolved Remarks | Reported At |
' Issue ID | Issue Title
' ============================================================================
' FIX APPLIED: .DrawingObject.Formula error fixed by using hidden cells
' and Worksheet_Calculate event for KPI updates
' ============================================================================

Sub BuildIssueDashboard()

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim wsDash As Worksheet
    Dim lo As ListObject
    Dim pc As PivotCache
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRng As Range

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wb = ThisWorkbook
    
    ' Check if Data sheet exists
    On Error Resume Next
    Set wsData = wb.Sheets("Data")
    On Error GoTo 0
    
    If wsData Is Nothing Then
        MsgBox "Please create a sheet named 'Data' with your issue tracking data!", vbCritical, "Error"
        Exit Sub
    End If

    ' STEP 1: Convert Data range to Excel Table
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Set dataRng = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))

    On Error Resume Next
    wsData.ListObjects("tblIssues").Unlist
    On Error GoTo 0

    Set lo = wsData.ListObjects.Add(xlSrcRange, dataRng, , xlYes)
    lo.Name = "tblIssues"
    lo.TableStyle = "TableStyleMedium2"

    ' STEP 2: Delete and recreate Pivot sheet
    On Error Resume Next
    wb.Sheets("Pivot").Delete
    On Error GoTo 0

    Set wsPivot = wb.Sheets.Add(After:=wsData)
    wsPivot.Name = "Pivot"

    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.ListObjects("tblIssues").Range)

    Call CreatePivot_IssueType(pc, wsPivot)
    Call CreatePivot_Severity(pc, wsPivot)
    Call CreatePivot_Status(pc, wsPivot)
    Call CreatePivot_Department(pc, wsPivot)
    Call CreatePivot_Resolver(pc, wsPivot)
    Call CreatePivot_Location(pc, wsPivot)
    Call CreatePivot_TrendMonthly(pc, wsPivot)

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

    ' Column widths setup
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

    ' STEP 4: Create Header Banner
    Call CreateHeaderBanner(wsDash)

    ' STEP 5: Setup KPI Hidden Cells (FIX for DrawingObject error)
    Call SetupKPIFormulas(wsDash, wsData)

    ' STEP 6: Create KPI Cards
    Call CreateKPICards(wsDash)

    ' STEP 7: Create Charts
    Call CreateCharts(wsDash, wsPivot)

    ' STEP 8: Create Slicers
    Call CreateSlicers(wb, wsDash, wsPivot)

    wsDash.Activate
    wsDash.Range("A1").Select

    ' STEP 9: Restore settings and apply zoom
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

    ' Update KPI text values
    Call UpdateKPIText(wsDash)

    MsgBox "Issue Tracking Dashboard built successfully!" & vbCrLf & _
           "Total Issues Loaded: " & (lastRow - 1), vbInformation, "Dashboard Ready"

End Sub

' ============================================================================
' HEADER BANNER
' ============================================================================
Private Sub CreateHeaderBanner(wsDash As Worksheet)

    Dim hdrLeft As Single, hdrTop As Single, hdrWidth As Single, hdrHeight As Single
    hdrLeft = wsDash.Range("A1").Left
    hdrTop = wsDash.Range("A1").Top
    hdrWidth = wsDash.Range("A1:X1").Width
    hdrHeight = wsDash.Range("A1:A3").Height

    Dim hdrShape As Shape
    Set hdrShape = wsDash.Shapes.AddShape(msoShapeRectangle, _
        hdrLeft, hdrTop, hdrWidth, hdrHeight)
    
    With hdrShape
        .Fill.ForeColor.RGB = RGB(21, 67, 96)
        .Fill.Transparency = 0
        .Line.Visible = msoFalse
        .Name = "HeaderBanner"
    End With
    
    With hdrShape.TextFrame
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .Characters.Text = "Issue Tracking Analytics Dashboard"
        .Characters.Font.Bold = True
        .Characters.Font.Color = RGB(255, 255, 255)
        .Characters.Font.Size = 32
        .Characters.Font.Name = "Calibri"
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With

End Sub

' ============================================================================
' SETUP KPI FORMULAS IN HIDDEN CELLS (FIX for DrawingObject error)
' ============================================================================
Private Sub SetupKPIFormulas(wsDash As Worksheet, wsData As Worksheet)

    Const KPI_ROW As Long = 200
    
    ' Total Issues
    wsDash.Cells(KPI_ROW, 1).Formula = "=COUNTA(Data!A2:A10000)"
    
    ' Open/Pending Issues
    wsDash.Cells(KPI_ROW, 2).Formula = _
        "=COUNTIF(Data!C2:C10000,""Open"")+COUNTIF(Data!C2:C10000,""Pending"")+COUNTIF(Data!C2:C10000,""In Progress"")"
    
    ' Resolved Issues
    wsDash.Cells(KPI_ROW, 3).Formula = _
        "=COUNTIF(Data!C2:C10000,""Resolved"")+COUNTIF(Data!C2:C10000,""Closed"")+COUNTIF(Data!C2:C10000,""Completed"")"
    
    ' Critical/High Severity Issues
    wsDash.Cells(KPI_ROW, 4).Formula = _
        "=COUNTIF(Data!B2:B10000,""Critical"")+COUNTIF(Data!B2:B10000,""High"")"
    
    ' Resolution Rate %
    wsDash.Cells(KPI_ROW, 5).Formula = _
        "=IFERROR((" & wsDash.Cells(KPI_ROW, 3).Address & "/" & wsDash.Cells(KPI_ROW, 1).Address & "),0)"
    
    ' Unique Reporters
    wsDash.Cells(KPI_ROW, 6).Formula = _
        "=SUMPRODUCT((Data!I2:I10000<>"""")/COUNTIF(Data!I2:I10000,Data!I2:I10000&""""))"
    
    ' Hide the row
    wsDash.Rows(KPI_ROW).Hidden = True
    wsDash.Rows(KPI_ROW + 1).Hidden = True

End Sub

' ============================================================================
' KPI CARDS
' ============================================================================
Private Sub CreateKPICards(wsDash As Worksheet)

    Dim kpiTitles(5) As String
    kpiTitles(0) = "Total Issues"
    kpiTitles(1) = "Open Issues"
    kpiTitles(2) = "Resolved"
    kpiTitles(3) = "Critical/High"
    kpiTitles(4) = "Resolution %"
    kpiTitles(5) = "Total Reporters"

    Dim accentColors(5) As Long
    accentColors(0) = RGB(37, 99, 235)
    accentColors(1) = RGB(245, 158, 11)
    accentColors(2) = RGB(5, 150, 105)
    accentColors(3) = RGB(220, 38, 38)
    accentColors(4) = RGB(124, 58, 237)
    accentColors(5) = RGB(21, 67, 96)

    Dim areaLeft As Single, areaWidth As Single, cardGap As Single
    Dim cardW As Single, cardH As Single, accentH As Single, cardTop As Single
    
    areaLeft = wsDash.Range("A1").Left + 5
    areaWidth = wsDash.Range("A1:T1").Width - 10
    cardGap = 8
    cardW = Int((areaWidth - 5 * cardGap) / 6)
    cardH = 85
    accentH = 12
    cardTop = wsDash.Range("A4").Top + 4

    Dim k As Integer
    Const KPI_ROW As Long = 200
    
    For k = 0 To 5
        Dim cLeft As Single
        cLeft = areaLeft + k * (cardW + cardGap)

        ' Main card shape
        Dim cardShape As Shape
        Set cardShape = wsDash.Shapes.AddShape(msoShapeRoundedRectangle, _
            cLeft, cardTop, cardW, cardH)
        
        With cardShape
            .Name = "KPI_Card_" & k
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Fill.Transparency = 0
            .Line.ForeColor.RGB = RGB(218, 224, 235)
            .Line.Weight = 0.75
            .Shadow.Type = msoShadow21
            .Shadow.Transparency = 0.75
            .Shadow.OffsetX = 1
            .Shadow.OffsetY = 2
            .Shadow.Size = 100
            .Shadow.Blur = 4
        End With

        ' Accent bar
        Dim acBar As Shape
        Set acBar = wsDash.Shapes.AddShape(msoShapeRectangle, _
            cLeft, cardTop, cardW, accentH)
        
        With acBar
            .Name = "KPI_Accent_" & k
            .Fill.ForeColor.RGB = accentColors(k)
            .Fill.Transparency = 0
            .Line.Visible = msoFalse
        End With

        ' Title textbox
        Dim titleBox As Shape
        Set titleBox = wsDash.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            cLeft + 6, cardTop + accentH + 5, cardW - 12, 22)
        
        With titleBox
            .Name = "KPI_Title_" & k
            .Line.Visible = msoFalse
            .Fill.Visible = msoFalse
            With .TextFrame
                .Characters.Text = kpiTitles(k)
                .Characters.Font.Name = "Calibri"
                .Characters.Font.Size = 11
                .Characters.Font.Bold = True
                .Characters.Font.Color = RGB(60, 60, 60)
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
        End With

        ' Value textbox - FIX: Use static text, updated by UpdateKPIText sub
        Dim valTop As Single, valH As Single
        valTop = cardTop + accentH + 30
        valH = cardH - accentH - 33
        
        Dim valBox As Shape
        Set valBox = wsDash.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            cLeft + 4, valTop, cardW - 8, valH)
        
        With valBox
            .Name = "KPI_Value_" & k
            .Line.Visible = msoFalse
            .Fill.Visible = msoFalse
            With .TextFrame
                .Characters.Text = "0"
                .Characters.Font.Name = "Calibri"
                .Characters.Font.Size = 22
                .Characters.Font.Bold = True
                .Characters.Font.Color = RGB(15, 32, 65)
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
        End With

    Next k

End Sub

' ============================================================================
' UPDATE KPI TEXT FROM HIDDEN CELLS (FIX for DrawingObject error)
' ============================================================================
Public Sub UpdateKPIText(wsDash As Worksheet)
    
    On Error Resume Next
    
    Const KPI_ROW As Long = 200
    Dim k As Integer
    Dim cellValue As Variant
    Dim displayText As String
    
    For k = 0 To 5
        cellValue = wsDash.Cells(KPI_ROW, k + 1).Value
        
        ' Format based on KPI type
        Select Case k
            Case 4 ' Resolution Rate - format as percentage
                displayText = Format(cellValue, "0.0%")
            Case Else ' Other - format as number
                displayText = Format(cellValue, "#,##0")
        End Select
        
        ' Update textbox text directly (NO DrawingObject.Formula needed!)
        wsDash.Shapes("KPI_Value_" & k).TextFrame.Characters.Text = displayText
    Next k
    
    On Error GoTo 0
End Sub

' ============================================================================
' PIVOT TABLES
' ============================================================================
Private Sub CreatePivot_IssueType(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A2"), _
        TableName:="pvtIssueType")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Issue Type").Orientation = xlRowField
        .PivotFields("Issue Type").Position = 1

        With .PivotFields("Issue ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count by Type"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_Severity(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A20"), _
        TableName:="pvtSeverity")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Severity").Orientation = xlRowField
        .PivotFields("Severity").Position = 1

        With .PivotFields("Issue ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count by Severity"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_Status(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A35"), _
        TableName:="pvtStatus")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Current Status").Orientation = xlRowField
        .PivotFields("Current Status").Position = 1

        With .PivotFields("Issue ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count by Status"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_Department(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("F2"), _
        TableName:="pvtDepartment")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Reporter Department").Orientation = xlRowField
        .PivotFields("Reporter Department").Position = 1

        With .PivotFields("Issue ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Issues by Department"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_Resolver(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    Dim pf As PivotField
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("F25"), _
        TableName:="pvtResolver")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        Set pf = .PivotFields("Resolver")
        pf.Orientation = xlRowField
        pf.Position = 1

        With .PivotFields("Issue ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Issues per Resolver"
            .NumberFormat = "#,##0"
        End With

        ' Show top 10 resolvers
        On Error Resume Next
        pf.AutoShow 1, 1, 10, "Issues per Resolver"
        On Error GoTo 0

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_Location(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("K2"), _
        TableName:="pvtLocation")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Issue Location").Orientation = xlRowField
        .PivotFields("Issue Location").Position = 1

        With .PivotFields("Issue ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Issues by Location"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_TrendMonthly(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("K25"), _
        TableName:="pvtTrend")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Reported At").Orientation = xlRowField
        .PivotFields("Reported At").Position = 1
        
        ' Group by Month
        On Error Resume Next
        .PivotFields("Reported At").LabelRange.Group Start:=True, End:=True, _
            Periods:=Array(False, False, False, False, True, False, True)
        On Error GoTo 0

        With .PivotFields("Issue ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Monthly Issues"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

' ============================================================================
' CHARTS
' ============================================================================
Private Sub CreateCharts(wsDash As Worksheet, wsPivot As Worksheet)

    Const CHART_BG As Long = 16777215
    Const CHART_NAVY As Long = 984097

    Dim startL As Single, contentW As Single, cGap As Single
    startL = wsDash.Range("A1").Left + 5
    contentW = wsDash.Range("A1:T1").Width - 10
    cGap = 8

    Dim rw1Top As Single, rw1H As Single, rw2Top As Single, rw2H As Single
    rw1Top = wsDash.Range("A12").Top + 3
    rw1H = wsDash.Range("A12:A26").Height - 6
    rw2Top = wsDash.Range("A27").Top + 3
    rw2H = wsDash.Range("A27:A42").Height - 6

    ' Chart widths
    Dim cht1W As Single, cht2W As Single, cht3W As Single
    cht1W = Int(contentW * 0.5)
    cht2W = contentW - cht1W - cGap
    cht3W = Int((contentW - 2 * cGap) / 3)

    ' Chart positions
    Dim c1L As Single, c2L As Single, c3L As Single, c4L As Single, c5L As Single
    c1L = startL
    c2L = c1L + cht1W + cGap
    c3L = startL
    c4L = c3L + cht3W + cGap
    c5L = c4L + cht3W + cGap

    ' CHART 1: Severity Distribution (Donut)
    Call MakeDonutChart(wsDash, wsPivot, "chtSeverity", _
        "Severity Distribution", "A20", c1L, rw1Top, cht1W, rw1H)

    ' CHART 2: Status Breakdown (Bar)
    Call MakeBarChart(wsDash, wsPivot, "chtStatus", _
        "Current Status Breakdown", "A35", c2L, rw1Top, cht2W, rw1H, _
        RGB(37, 99, 235))

    ' CHART 3: Issues by Department (Column)
    Call MakeColumnChart(wsDash, wsPivot, "chtDepartment", _
        "Issues by Department", "F2", c3L, rw2Top, cht3W, rw2H, _
        RGB(5, 150, 105))

    ' CHART 4: Top 10 Resolvers (Bar)
    Call MakeBarChart(wsDash, wsPivot, "chtResolver", _
        "Top 10 Resolvers", "F25", c4L, rw2Top, cht3W, rw2H, _
        RGB(124, 58, 237))

    ' CHART 5: Issues by Location (Column)
    Call MakeColumnChart(wsDash, wsPivot, "chtLocation", _
        "Issues by Location", "K2", c5L, rw2Top, cht3W, rw2H, _
        RGB(245, 158, 11))

End Sub

Private Sub MakeDonutChart(wsDash As Worksheet, wsPivot As Worksheet, _
    chtName As String, chtTitle As String, srcRange As String, _
    L As Single, T As Single, W As Single, H As Single)
    
    On Error Resume Next
    wsDash.ChartObjects(chtName).Delete
    On Error GoTo 0
    
    Dim cht As ChartObject
    Set cht = wsDash.ChartObjects.Add(L, T, W, H)
    cht.Name = chtName
    
    With cht.Chart
        .ChartType = xlDoughnut
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range(srcRange).PivotTable.TableRange1
        On Error GoTo 0
        
        .HasTitle = True
        .ChartTitle.Text = chtTitle
        .ChartTitle.Font.Size = 11
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(15, 32, 65)
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        .Legend.Font.Size = 9
        
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color = RGB(255, 255, 255)
        .PlotArea.Interior.Color = RGB(255, 255, 255)
        
        On Error Resume Next
        .SeriesCollection(1).DoughnutHoleSize = 50
        
        Dim dColors(3) As Long
        dColors(0) = RGB(220, 38, 38)
        dColors(1) = RGB(245, 158, 11)
        dColors(2) = RGB(37, 99, 235)
        dColors(3) = RGB(5, 150, 105)
        
        Dim p As Integer
        For p = 1 To .SeriesCollection(1).Points.Count
            If p <= 4 Then
                .SeriesCollection(1).Points(p).Format.Fill.ForeColor.RGB = dColors(p - 1)
            End If
        Next p
        
        With .SeriesCollection(1)
            .HasDataLabels = True
            .DataLabels.ShowCategoryName = True
            .DataLabels.ShowPercentage = True
            .DataLabels.ShowValue = False
            .DataLabels.NumberFormat = "0%"
            .DataLabels.Font.Size = 9
            .DataLabels.Font.Bold = True
            .DataLabels.Font.Color = RGB(255, 255, 255)
        End With
        
        .ShowAllFieldButtons = False
        On Error GoTo 0
    End With
End Sub

Private Sub MakeColumnChart(wsDash As Worksheet, wsPivot As Worksheet, _
    chtName As String, chtTitle As String, srcRange As String, _
    L As Single, T As Single, W As Single, H As Single, barColor As Long)
    
    On Error Resume Next
    wsDash.ChartObjects(chtName).Delete
    On Error GoTo 0
    
    Dim cht As ChartObject
    Set cht = wsDash.ChartObjects.Add(L, T, W, H)
    cht.Name = chtName
    
    With cht.Chart
        .ChartType = xlColumnClustered
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range(srcRange).PivotTable.TableRange1
        On Error GoTo 0
        
        .HasTitle = True
        .ChartTitle.Text = chtTitle
        .ChartTitle.Font.Size = 11
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(15, 32, 65)
        
        .HasLegend = False
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color = RGB(255, 255, 255)
        .PlotArea.Interior.Color = RGB(250, 252, 255)
        
        On Error Resume Next
        With .SeriesCollection(1)
            .Format.Fill.ForeColor.RGB = barColor
            .GapWidth = 50
            .HasDataLabels = True
            .DataLabels.NumberFormat = "#,##0"
            .DataLabels.Font.Size = 8
            .DataLabels.Font.Bold = True
            .DataLabels.Font.Color = RGB(15, 32, 65)
            .DataLabels.Position = xlLabelPositionOutsideEnd
        End With
        
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 234, 242)
        .ShowAllFieldButtons = False
        On Error GoTo 0
    End With
End Sub

Private Sub MakeBarChart(wsDash As Worksheet, wsPivot As Worksheet, _
    chtName As String, chtTitle As String, srcRange As String, _
    L As Single, T As Single, W As Single, H As Single, barColor As Long)
    
    On Error Resume Next
    wsDash.ChartObjects(chtName).Delete
    On Error GoTo 0
    
    Dim cht As ChartObject
    Set cht = wsDash.ChartObjects.Add(L, T, W, H)
    cht.Name = chtName
    
    With cht.Chart
        .ChartType = xlBarClustered
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range(srcRange).PivotTable.TableRange1
        On Error GoTo 0
        
        .HasTitle = True
        .ChartTitle.Text = chtTitle
        .ChartTitle.Font.Size = 11
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(15, 32, 65)
        
        .HasLegend = False
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color = RGB(255, 255, 255)
        .PlotArea.Interior.Color = RGB(250, 252, 255)
        
        On Error Resume Next
        With .SeriesCollection(1)
            .Format.Fill.ForeColor.RGB = barColor
            .GapWidth = 50
            .HasDataLabels = True
            .DataLabels.NumberFormat = "#,##0"
            .DataLabels.Font.Size = 8
            .DataLabels.Font.Bold = True
            .DataLabels.Font.Color = RGB(15, 32, 65)
            .DataLabels.Position = xlLabelPositionOutsideEnd
        End With
        
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 234, 242)
        .ShowAllFieldButtons = False
        On Error GoTo 0
    End With
End Sub

' ============================================================================
' SLICERS
' ============================================================================
Private Sub CreateSlicers(wb As Workbook, wsDash As Worksheet, wsPivot As Worksheet)

    Dim panelLeft As Single, panelTop As Single
    Dim panelWidth As Single, panelHeight As Single
    
    panelLeft = wsDash.Range("V4").Left
    panelTop = wsDash.Range("V4").Top
    panelWidth = wsDash.Range("V4:X4").Width
    panelHeight = wsDash.Range("V4:V42").Height

    ' Panel background
    Dim panelBG As Shape
    Set panelBG = wsDash.Shapes.AddShape(msoShapeRectangle, _
        panelLeft, panelTop, panelWidth, panelHeight)
    
    With panelBG
        .Name = "SlicerPanel_BG"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Transparency = 0
        .Line.ForeColor.RGB = RGB(210, 218, 232)
        .Line.Weight = 0.75
        .ZOrder msoSendToBack
    End With

    ' Slicer fields
    Dim slicerFields(3) As String
    slicerFields(0) = "Issue Type"
    slicerFields(1) = "Severity"
    slicerFields(2) = "Current Status"
    slicerFields(3) = "Reporter Department"

    Dim slicerNames(3) As String
    slicerNames(0) = "slcIssueType"
    slicerNames(1) = "slcSeverity"
    slicerNames(2) = "slcStatus"
    slicerNames(3) = "slcDept"

    Dim slicerH(3) As Single
    slicerH(0) = 90
    slicerH(1) = 90
    slicerH(2) = 110
    slicerH(3) = 130

    ' Pivot tables to connect
    Dim ptNames(6) As String
    ptNames(0) = "pvtIssueType"
    ptNames(1) = "pvtSeverity"
    ptNames(2) = "pvtStatus"
    ptNames(3) = "pvtDepartment"
    ptNames(4) = "pvtResolver"
    ptNames(5) = "pvtLocation"
    ptNames(6) = "pvtTrend"

    ' Delete existing slicer caches
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

    ' Create new slicers
    Dim slicerW As Single, slicerLeft As Single, currentTop As Single
    slicerW = panelWidth - 8
    slicerLeft = panelLeft + 4
    currentTop = panelTop + 4

    Dim f As Integer
    For f = 0 To 3
        Dim sc As SlicerCache
        Dim sl As Slicer
        Dim firstPT As PivotTable

        On Error Resume Next
        Set firstPT = wsPivot.PivotTables("pvtIssueType")
        Set sc = wb.SlicerCaches.Add2(firstPT, slicerFields(f), slicerNames(f))

        ' Connect to all pivots
        Dim pn As Integer
        For pn = 0 To 6
            On Error Resume Next
            sc.PivotTables.AddPivotTable wsPivot.PivotTables(ptNames(pn))
            On Error GoTo 0
        Next pn

        Set sl = sc.Slicers.Add(wsDash, , slicerNames(f) & "_visual", _
            slicerFields(f), currentTop, slicerLeft, slicerW, slicerH(f))

        With sl
            .Style = "SlicerStyleDark4"
            .Left = slicerLeft
            .Top = currentTop
            .Width = slicerW
            .Height = slicerH(f)
            .NumberOfColumns = 1
            .RowHeight = 18
        End With
        On Error GoTo 0

        currentTop = currentTop + slicerH(f) + 8
    Next f

End Sub

' ============================================================================
' REFRESH KPI - Manual refresh button function
' ============================================================================
Public Sub RefreshDashboard()
    On Error Resume Next
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    
    If Not wsDash Is Nothing Then
        wsDash.Calculate
        Call UpdateKPIText(wsDash)
    End If
    On Error GoTo 0
End Sub
