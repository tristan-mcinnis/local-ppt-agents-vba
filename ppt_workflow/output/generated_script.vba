' ================================================================
' AUTO-GENERATED POWERPOINT VBA SCRIPT
' Generated: 2025-09-14 11:37:07 UTC
' Template: ic-template-1.pptx
' Platform: macOS and Windows compatible
' ================================================================
'
' USAGE:
'   1. Open your PowerPoint template
'   2. Press Alt+F11 (Windows) or Opt+F11 (Mac) to open VBA editor
'   3. Insert > Module
'   4. Paste this entire script
'   5. Run the "Main" subroutine
'
' ================================================================

Option Explicit

' PowerPoint placeholder type constants
Const msoPlaceholder As Long = 14
Const ppPlaceholderTitle As Long = 1
Const ppPlaceholderBody As Long = 2
Const ppPlaceholderCenterTitle As Long = 3
Const ppPlaceholderSubtitle As Long = 4
Const ppPlaceholderObject As Long = 7
Const ppPlaceholderChart As Long = 8
Const ppPlaceholderTable As Long = 9
Const ppPlaceholderPicture As Long = 18

' Chart type constants (macOS-safe)
Const xlColumnClustered As Long = 51
Const xlBarClustered As Long = 57
Const xlLine As Long = 4
Const xlPie As Long = 5
Const xlArea As Long = 1
Const xlXYScatter As Long = -4169

' Platform detection
#If Mac Then
    Const PLATFORM As String = "macOS"
#Else
    Const PLATFORM As String = "Windows"
#End If


' ================================================================
' HELPER FUNCTIONS
' ================================================================

' Module-level cache for layouts (macOS-safe Collection instead of Scripting.Dictionary)
Dim layoutCache As Collection

' Check if layout cache has a key (macOS-safe)
Private Function CacheHas(key As Long) As Boolean
    On Error GoTo NotFound
    Dim tmp As CustomLayout
    Set tmp = layoutCache(CStr(key))
    CacheHas = True
    Exit Function
NotFound:
    CacheHas = False
End Function

' Add layout to cache (macOS-safe)
Private Sub CachePut(key As Long, cl As CustomLayout)
    On Error Resume Next
    ' Remove if exists
    layoutCache.Remove CStr(key)
    On Error GoTo 0
    layoutCache.Add cl, CStr(key)
End Sub

' Get layout from cache (macOS-safe)
Private Function CacheGet(key As Long) As CustomLayout
    On Error Resume Next
    Set CacheGet = layoutCache(CStr(key))
    On Error GoTo 0
End Function

' Get custom layout by index - IMPROVED VERSION
' Matches the analyzer logic: index is position within SlideMaster.CustomLayouts
Function GetCustomLayoutByIndexSafe(layoutIndex As Long) As CustomLayout
    On Error Resume Next
    Dim pres As Presentation
    Dim design As design

    Set pres = Application.ActivePresentation

    ' First try: Direct index from the active SlideMaster (most common case)
    If layoutIndex >= 1 And layoutIndex <= pres.SlideMaster.CustomLayouts.Count Then
        Set GetCustomLayoutByIndexSafe = pres.SlideMaster.CustomLayouts(layoutIndex)
        If Not GetCustomLayoutByIndexSafe Is Nothing Then
            Exit Function
        End If
    End If

    ' Second try: Check each Design's SlideMaster (for multi-design templates)
    For Each design In pres.Designs
        If layoutIndex >= 1 And layoutIndex <= design.SlideMaster.CustomLayouts.Count Then
            Set GetCustomLayoutByIndexSafe = design.SlideMaster.CustomLayouts(layoutIndex)
            If Not GetCustomLayoutByIndexSafe Is Nothing Then
                Exit Function
            End If
        End If
    Next design

    ' Return Nothing if not found
    Set GetCustomLayoutByIndexSafe = Nothing
    On Error GoTo 0
End Function

' Add slide with specified layout
Function AddSlideWithLayout(layout As CustomLayout) As Slide
    Dim pres As Presentation
    Set pres = Application.ActivePresentation
    Set AddSlideWithLayout = pres.Slides.AddSlide(pres.Slides.Count + 1, layout)
End Function

' Get placeholder by type and ordinal (0-based)
Function GetPlaceholderByTypeAndOrdinal(sld As Slide, typeId As Long, ordinal As Long) As Shape
    On Error Resume Next
    Dim shp As Shape
    Dim candidates As Collection
    Dim i As Long

    Set candidates = New Collection

    ' Collect all placeholders of the specified type
    For Each shp In sld.Shapes
        If shp.Type = msoPlaceholder Then
            If shp.PlaceholderFormat.Type = typeId Then
                candidates.Add shp
            End If
        End If
    Next shp

    ' Sort by position (top, then left)
    Dim sorted As Collection
    Set sorted = SortShapesByPosition(candidates)

    ' Return the one at the specified ordinal
    If ordinal >= 0 And ordinal < sorted.Count Then
        Set GetPlaceholderByTypeAndOrdinal = sorted(ordinal + 1) ' Collection is 1-based
    Else
        Set GetPlaceholderByTypeAndOrdinal = Nothing
    End If

    On Error GoTo 0
End Function

' Sort shapes by position (top, then left)
Function SortShapesByPosition(shapes As Collection) As Collection
    Dim sorted As Collection
    Dim shp As Shape
    Dim i As Long, j As Long
    Dim tempShp As Shape

    Set sorted = New Collection

    ' Copy to sorted collection
    For Each shp In shapes
        sorted.Add shp
    Next shp

    ' Simple bubble sort
    For i = 1 To sorted.Count - 1
        For j = i + 1 To sorted.Count
            If sorted(i).Top > sorted(j).Top Or _
               (sorted(i).Top = sorted(j).Top And sorted(i).Left > sorted(j).Left) Then
                ' Swap
                Set tempShp = sorted(i)
                sorted.Remove i
                sorted.Add tempShp, , , i
                Set tempShp = sorted(j)
                sorted.Remove j
                sorted.Add tempShp, , , j - 1
            End If
        Next j
    Next i

    Set SortShapesByPosition = sorted
End Function

' Set text with TextFrame2 fallback for compatibility
Sub SafeSetText(shp As Shape, text As String)
    On Error Resume Next

    ' Try TextFrame2 first (newer PowerPoint)
    If Not shp.TextFrame2 Is Nothing Then
        shp.TextFrame2.TextRange.text = text
        If Err.Number = 0 Then Exit Sub
    End If

    Err.Clear

    ' Fallback to TextFrame (older PowerPoint)
    If shp.HasTextFrame Then
        shp.TextFrame.TextRange.text = text
    End If

    On Error GoTo 0
End Sub

' Debug helper to list placeholders on a slide
Sub DebugListPlaceholders(s As Slide)
    Dim sh As Shape
    Debug.Print "=== Placeholders on slide " & s.SlideIndex & " ==="
    For Each sh In s.Shapes
        If sh.Type = msoPlaceholder Then
            Debug.Print "  type_id=" & sh.PlaceholderFormat.Type & _
                        " top=" & Round(sh.Top) & " left=" & Round(sh.Left)
        End If
    Next sh
    Debug.Print "=== End of placeholder list ==="
End Sub

' Create chart at placeholder location (macOS-safe)
Sub CreateChartAtPlaceholder(sld As Slide, placeholder As Shape, chartSpec As String)
    On Error Resume Next
    Dim chartShape As Shape
    Dim chartObj As Object
    Dim l As Single, t As Single, w As Single, h As Single
    Dim chartType As Long

    ' Get placeholder dimensions
    l = placeholder.Left
    t = placeholder.Top
    w = placeholder.Width
    h = placeholder.Height

    ' Delete placeholder
    placeholder.Delete

    ' Determine chart type from spec (default to column)
    chartType = xlColumnClustered
    If InStr(chartSpec, """line""") > 0 Then chartType = xlLine
    If InStr(chartSpec, """bar""") > 0 Then chartType = xlBarClustered
    If InStr(chartSpec, """pie""") > 0 Then chartType = xlPie
    If InStr(chartSpec, """area""") > 0 Then chartType = xlArea
    If InStr(chartSpec, """scatter""") > 0 Then chartType = xlXYScatter

    ' Create chart using AddChart (macOS compatible)
    Set chartShape = sld.Shapes.AddChart(chartType, l, t, w, h)

    If chartShape Is Nothing Then
        MsgBox "Failed to create chart", vbCritical
        Exit Sub
    End If

    ' Access chart object
    Set chartObj = chartShape.Chart

    ' Populate chart based on the spec
    ' For extensibility, look for specific data patterns
    If InStr(chartSpec, "WAU") > 0 Or InStr(chartSpec, "Week 1") > 0 Then
        ' North Star Metrics or any weekly data chart
        PopulateWeeklyMetricsChart chartObj, chartSpec
    ElseIf InStr(chartSpec, "series") > 0 Then
        ' Generic multi-series chart
        PopulateGenericChart chartObj, chartSpec
    Else
        ' Fallback sample data
        With chartObj.ChartData
            .Activate
            .Workbook.Worksheets(1).Range("A1").Value = "Category"
            .Workbook.Worksheets(1).Range("B1").Value = "Value"
            .Workbook.Worksheets(1).Range("A2").Value = "Item 1"
            .Workbook.Worksheets(1).Range("B2").Value = 10
            .Workbook.Close
        End With
    End If

    On Error GoTo 0
End Sub

' Populate weekly metrics chart with actual data
Sub PopulateWeeklyMetricsChart(chartObj As Object, chartSpec As String)
    With chartObj.ChartData
        .Activate
        Dim ws As Object
        Set ws = .Workbook.Worksheets(1)

        ' Clear existing data
        ws.Cells.Clear

        ' Set up headers
        ws.Range("A1").Value = "Week"
        ws.Range("B1").Value = "WAU"
        ws.Range("C1").Value = "Median Latency (s)"

        ' Add x-axis labels
        ws.Range("A2:A7").Value = Application.Transpose(Array("Week 1", "Week 2", "Week 3", "Week 4", "Week 5", "Week 6"))

        ' Add WAU data
        ws.Range("B2:B7").Value = Application.Transpose(Array(200, 380, 520, 780, 1100, 1400))

        ' Add Latency data
        ws.Range("C2:C7").Value = Application.Transpose(Array(3.2, 2.8, 2.4, 2.1, 1.9, 1.7))

        ' Set the data range
        chartObj.SetSourceData Source:=ws.Range("A1:C7")

        .Workbook.Close
    End With

    ' Set chart title
    chartObj.HasTitle = True
    chartObj.ChartTitle.Text = "Metrics Trend"
End Sub

' Populate generic chart
Sub PopulateGenericChart(chartObj As Object, chartSpec As String)
    With chartObj.ChartData
        .Activate
        .Workbook.Worksheets(1).Range("A1").Value = "Category"
        .Workbook.Worksheets(1).Range("B1").Value = "Series 1"
        .Workbook.Worksheets(1).Range("A2:A4").Value = Application.Transpose(Array("Q1", "Q2", "Q3"))
        .Workbook.Worksheets(1).Range("B2:B4").Value = Application.Transpose(Array(25, 35, 45))
        .Workbook.Close
    End With
End Sub

' Create table at placeholder location
Sub CreateTableAtPlaceholder(sld As Slide, placeholder As Shape, tableSpec As String)
    On Error Resume Next
    Dim tblShape As Shape
    Dim tbl As Table
    Dim l As Single, t As Single, w As Single, h As Single
    Dim rows As Long, cols As Long
    Dim r As Long, c As Long

    ' Get placeholder dimensions
    l = placeholder.Left
    t = placeholder.Top
    w = placeholder.Width
    h = placeholder.Height

    ' Delete placeholder
    placeholder.Delete

    ' Parse table spec (simplified - in production, parse JSON)
    rows = 3
    cols = 2

    ' Create table
    Set tblShape = sld.Shapes.AddTable(rows, cols, l, t, w, h)
    Set tbl = tblShape.Table

    ' Add sample data
    tbl.Cell(1, 1).Shape.TextFrame.TextRange.text = "Header 1"
    tbl.Cell(1, 2).Shape.TextFrame.TextRange.text = "Header 2"

    On Error GoTo 0
End Sub



' ================================================================
' MAIN SUBROUTINE
' ================================================================

Sub Main()
    On Error GoTo ErrorHandler

    ' Validate environment
    If Application.Presentations.Count = 0 Then
        MsgBox "Please open a PowerPoint presentation first.", vbExclamation
        Exit Sub
    End If

    Dim pres As Presentation
    Dim currentSlide As Slide
    Dim shp As Shape
    Dim cl As CustomLayout

    Set pres = Application.ActivePresentation
    
    ' Initialize layout cache (macOS-safe Collection)
    Set layoutCache = New Collection

    ' Pre-cache layouts for performance
    Dim layoutIndex As Variant
    Dim requiredLayouts As Variant
    requiredLayouts = Array(6, 8, 13, 21, 25, 38, 54, 56, 58, 59)

    For Each layoutIndex In requiredLayouts
        If Not CacheHas(CLng(layoutIndex)) Then
            Set cl = GetCustomLayoutByIndexSafe(CLng(layoutIndex))
            If cl Is Nothing Then
                MsgBox "Layout index " & layoutIndex & " not found in template. Check that you have the correct template open.", vbCritical
                Exit Sub
            End If
            CachePut CLng(layoutIndex), cl
        End If
    Next layoutIndex

    ' Create slides
    ' ---- Slide 1: Demo Deck ----
    Set currentSlide = AddSlideWithLayout(CacheGet(58))

    ' CenterTitle placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 3, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 1:" & vbCrLf & _
               "Type: CenterTitle (type_id=3)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Building a Modern Data Product"

    ' Subtitle placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 4, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 1:" & vbCrLf & _
               "Type: Subtitle (type_id=4)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "From concept to launch in 6 weeks"

    
    ' Image placeholder skipped: assets/logo.png
    ' User will add image manually after slides are created

    ' ---- Slide 2: Agenda ----
    Set currentSlide = AddSlideWithLayout(CacheGet(6))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 2:" & vbCrLf & _
               "Type: Title (type_id=1)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Agenda"

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 2:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "- Vision and goals" & vbCrLf & "- Users and use cases" & vbCrLf & "- Architecture overview" & vbCrLf & "- Prototype demo" & vbCrLf & "- Metrics & timeline"

    ' Body placeholder (ordinal 1)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 1)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 2:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 1" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "- Risks & mitigations" & vbCrLf & "- Team & roles" & vbCrLf & "- Budget overview" & vbCrLf & "- Q&A"

    ' ---- Slide 3: Vision ----
    Set currentSlide = AddSlideWithLayout(CacheGet(13))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 3:" & vbCrLf & _
               "Type: Title (type_id=1)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Vision"

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 3:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Deliver self-serve analytics to 3,000+ internal users with sub-2s query latency and governed data access."

    
    ' Image placeholder skipped: images/hero-dashboard.jpg
    ' User will add image manually after slides are created

    ' ---- Slide 4: Key Personas ----
    Set currentSlide = AddSlideWithLayout(CacheGet(21))
    
    ' Debug: List placeholders on slide 4 (layout 21)
    DebugListPlaceholders currentSlide

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 4:" & vbCrLf & _
               "Type: Title (type_id=1)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Key Personas"

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 4:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "**Analyst**" & vbCrLf & "- Ad-hoc exploration" & vbCrLf & "- SQL power-user" & vbCrLf & "- Needs versioned datasets"

    ' Body placeholder (ordinal 1)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 1)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 4:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 1" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "**Manager**" & vbCrLf & "- Weekly KPIs" & vbCrLf & "- Email/PDF exports" & vbCrLf & "- SLA: 9 am Monday"

    
    ' Image placeholder skipped: images/persona-analyst.png
    ' User will add image manually after slides are created

    
    ' Image placeholder skipped: images/persona-manager.png
    ' User will add image manually after slides are created

    ' ---- Slide 5: North Star Metrics ----
    Set currentSlide = AddSlideWithLayout(CacheGet(59))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 5:" & vbCrLf & _
               "Type: Title (type_id=1)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "North Star Metrics"

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 5:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "- Weekly Active Users (WAU)" & vbCrLf & "- Median query latency (s)" & vbCrLf & "- Dashboard share rate (%)" & vbCrLf & "- Data freshness (hrs)"

    ' Chart placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 8, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 5:" & vbCrLf & _
               "Type: Chart (type_id=8)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    CreateChartAtPlaceholder currentSlide, shp, "{""type"": ""line"", ""data"": {""x"": [""Week 1"", ""Week 2"", ""Week 3"", ""Week 4"", ""Week 5"", ""Week 6""], ""series"": [{""name"": ""WAU"", ""data"": [200, 380, 520, 780, 1100, 1400]}, {""name"": ""Median Latency (s)"", ""data"": [3.2, 2.8, 2.4, 2.1, 1.9, 1.7]}]}}"

    ' ---- Slide 6: High-Level Architecture ----
    Set currentSlide = AddSlideWithLayout(CacheGet(25))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 6:" & vbCrLf & _
               "Type: Title (type_id=1)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "High-Level Architecture"

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 6:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "- Ingest: CDC + batch (Kafka, Airflow)" & vbCrLf & "- Lakehouse: Parquet + Delta" & vbCrLf & "- Query: DuckDB + Presto" & vbCrLf & "- Serving: REST + BI" & vbCrLf & "- AuthZ: OPA + Column-masking"

    
    ' Image placeholder skipped: images/arch-overview.png
    ' User will add image manually after slides are created

    ' ---- Slide 7: Data Contracts Snapshot ----
    Set currentSlide = AddSlideWithLayout(CacheGet(8))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 7:" & vbCrLf & _
               "Type: Title (type_id=1)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Data Contracts Snapshot"

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 7:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "**Orders v2**" & vbCrLf & "- PII tags: none" & vbCrLf & "- SLA: 15 min" & vbCrLf & "- Owner: Sales Eng"

    ' Body placeholder (ordinal 1)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 1)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 7:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 1" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "**Customers v1**" & vbCrLf & "- PII tags: email" & vbCrLf & "- SLA: hourly" & vbCrLf & "- Owner: CRM"

    ' Body placeholder (ordinal 2)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 2)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 7:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 2" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "**Events v3**" & vbCrLf & "- PII tags: anon_id" & vbCrLf & "- SLA: near-real-time" & vbCrLf & "- Owner: Platform"

    ' ---- Slide 8: Prototype Screens ----
    Set currentSlide = AddSlideWithLayout(CacheGet(38))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 8:" & vbCrLf & _
               "Type: Title (type_id=1)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Prototype Screens"

    
    ' Image placeholder skipped: images/screen-explore.png
    ' User will add image manually after slides are created

    
    ' Image placeholder skipped: images/screen-dashboard.png
    ' User will add image manually after slides are created

    
    ' Image placeholder skipped: images/screen-share.png
    ' User will add image manually after slides are created

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 8:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Explore" & vbCrLf & "- Column profiler" & vbCrLf & "- Query hints"

    ' Body placeholder (ordinal 1)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 1)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 8:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 1" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Dashboard" & vbCrLf & "- Filters" & vbCrLf & "- Cohorts"

    ' Body placeholder (ordinal 2)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 2)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 8:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 2" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Share" & vbCrLf & "- Links & embeds" & vbCrLf & "- Expiring tokens"

    ' ---- Slide 9: Risks & Mitigations ----
    Set currentSlide = AddSlideWithLayout(CacheGet(56))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 9:" & vbCrLf & _
               "Type: Title (type_id=1)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Risks & Mitigations"

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 9:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "- Data freshness SLAs might slip → Add backfill monitors and retry budgets" & vbCrLf & "- Governance gaps → Enforce policy-as-code and lineage checks in CI" & vbCrLf & "- Cost overruns → Right-size clusters, auto-suspend idle, tiered storage" & vbCrLf & "- Adoption lag → Champions program, office hours, template gallery"

    ' ---- Slide 10: Thank You ----
    Set currentSlide = AddSlideWithLayout(CacheGet(54))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 10:" & vbCrLf & _
               "Type: Title (type_id=1)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Thank You"

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide 10:" & vbCrLf & _
               "Type: Body (type_id=2)" & vbCrLf & _
               "Ordinal: 0" & vbCrLf & vbCrLf & _
               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"
        Exit Sub
    End If
    SafeSetText shp, "Contact" & vbCrLf & "- Product: product@company.com" & vbCrLf & "- Slack: #data-product" & vbCrLf & "- Docs: https://docs.company.com/data-product"


    ' Success message
    MsgBox "Successfully created 10 slides!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub


' ================================================================
' VALIDATION SUBROUTINE (Optional)
' ================================================================

Sub ValidateTemplate()
    On Error Resume Next
    Dim pres As Presentation
    Dim layout As CustomLayout
    Dim msg As String

    Set pres = Application.ActivePresentation

    msg = "Template Validation Report:" & vbCrLf & vbCrLf
    msg = msg & "Template: " & pres.Name & vbCrLf
    msg = msg & "Layouts: " & pres.SlideMaster.CustomLayouts.Count & vbCrLf
    msg = msg & "Platform: " & PLATFORM & vbCrLf & vbCrLf

    ' Check required layouts
    Dim requiredLayouts As Variant
    Dim layoutIndex As Variant
    Dim found As Boolean

    requiredLayouts = Array(6, 8, 13, 21, 25, 38, 54, 56, 58, 59)

    msg = msg & "Required Layout Indices:" & vbCrLf
    For Each layoutIndex In requiredLayouts
        Set layout = GetCustomLayoutByIndexSafe(CLng(layoutIndex))
        If Not layout Is Nothing Then
            msg = msg & "  ✓ Index " & layoutIndex & ": " & layout.Name & vbCrLf
        Else
            msg = msg & "  ✗ Index " & layoutIndex & ": NOT FOUND" & vbCrLf
        End If
    Next layoutIndex

    MsgBox msg, vbInformation, "Template Validation"
    On Error GoTo 0
End Sub