' ================================================================
' AUTO-GENERATED POWERPOINT VBA SCRIPT
' Generated: 2025-09-17 01:10:43 UTC
' Template: Standard Template.pptx
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
Dim errorLog As Collection

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

' ----- Error Logging Helpers -----
Sub InitErrorLog()
    Set errorLog = New Collection
End Sub

Sub LogError(code As String, details As String)
    On Error Resume Next
    If errorLog Is Nothing Then Set errorLog = New Collection
    errorLog.Add code & ": " & details
    On Error GoTo 0
End Sub

Function ErrorsCount() As Long
    If errorLog Is Nothing Then
        ErrorsCount = 0
    Else
        ErrorsCount = errorLog.Count
    End If
End Function

Sub ShowErrors()
    If errorLog Is Nothing Or errorLog.Count = 0 Then Exit Sub
    Dim i As Long
    Dim msg As String
    msg = "Encountered " & errorLog.Count & " issue(s):" & vbCrLf & vbCrLf
    For i = 1 To errorLog.Count
        msg = msg & "- " & errorLog(i) & vbCrLf
        If i >= 12 Then
            msg = msg & "... (more omitted)" & vbCrLf
            Exit For
        End If
    Next i
    MsgBox msg, vbExclamation, "PowerPoint Script Issues"
End Sub

' ----- JSON Parsing Helpers (robust) -----
Private Sub Json_SkipWs(ByVal s As String, ByRef pos As Long)
    Dim ch As String
    Do While pos <= Len(s)
        ch = Mid$(s, pos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        pos = pos + 1
    Loop
End Sub

Private Function Json_ParseString(ByVal s As String, ByRef pos As Long) As String
    Dim res As String, ch As String, esc As Boolean
    Dim code As String, uVal As Integer
    ' Assumes current pos points at opening quote
    pos = pos + 1
    Do While pos <= Len(s)
        ch = Mid$(s, pos, 1)
        pos = pos + 1
        If esc Then
            Select Case ch
                Case """: res = res & ""
                Case "": res = res & ""
                Case "/": res = res & "/"
                Case "b": res = res & Chr$(8)
                Case "f": res = res & Chr$(12)
                Case "n": res = res & vbLf
                Case "r": res = res & vbCr
                Case "t": res = res & vbTab
                Case "u"
                    If pos + 3 <= Len(s) Then
                        code = Mid$(s, pos, 4)
                        On Error Resume Next
                        uVal = CInt("&H" & code)
                        If Err.Number = 0 Then
                            res = res & ChrW$(uVal)
                        Else
                            res = res & "?"
                            Err.Clear
                        End If
                        On Error GoTo 0
                        pos = pos + 4
                    End If
                Case Else
                    res = res & ch
            End Select
            esc = False
        ElseIf ch = "" Then
            esc = True
        ElseIf ch = """" Then
            Exit Do
        Else
            res = res & ch
        End If
    Loop
    Json_ParseString = res
End Function

Private Function Json_FindMatching(ByVal s As String, ByVal startPos As Long, _
                                   ByVal openCh As String, ByVal closeCh As String) As Long
    Dim i As Long, depth As Long, ch As String, esc As Boolean, inStr As Boolean
    depth = 0
    For i = startPos To Len(s)
        ch = Mid$(s, i, 1)
        If inStr Then
            If esc Then
                esc = False
            ElseIf ch = "" Then
                esc = True
            ElseIf ch = """" Then
                inStr = False
            End If
        Else
            If ch = """" Then
                inStr = True
            ElseIf ch = openCh Then
                depth = depth + 1
            ElseIf ch = closeCh Then
                depth = depth - 1
                If depth = 0 Then
                    Json_FindMatching = i
                    Exit Function
                End If
            End If
        End If
    Next i
    Json_FindMatching = 0
End Function

Private Function Json_FindKeyAtTopLevel(ByVal s As String, ByVal key As String) As Long
    ' Returns position of first character of value for key at top-level of object
    Dim i As Long, ch As String, esc As Boolean, inStr As Boolean, depth As Long
    Dim k As String, valPos As Long
    depth = 0
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If inStr Then
            If esc Then
                esc = False
            ElseIf ch = "" Then
                esc = True
            ElseIf ch = """" Then
                inStr = False
            End If
            i = i + 1
        Else
            Select Case ch
                Case """"  ' start string (potential key)
                    If depth = 1 Then
                        Dim savePos As Long
                        savePos = i
                        k = Json_ParseString(s, i)
                        Json_SkipWs s, i
                        If i <= Len(s) And Mid$(s, i, 1) = ":" Then
                            i = i + 1
                            Json_SkipWs s, i
                            If k = key Then
                                Json_FindKeyAtTopLevel = i
                                Exit Function
                            Else
                                ' skip the value to continue scanning
                                Dim c As String
                                c = Mid$(s, i, 1)
                                If c = """" Then
                                    Call Json_ParseString(s, i)
                                ElseIf c = "{" Then
                                    i = Json_FindMatching(s, i, "{", "}") + 1
                                ElseIf c = "[" Then
                                    i = Json_FindMatching(s, i, "[", "]") + 1
                                Else
                                    ' literal
                                    Do While i <= Len(s)
                                        c = Mid$(s, i, 1)
                                        If c = "," Or c = "}" Then Exit Do
                                        i = i + 1
                                    Loop
                                End If
                            End If
                        End If
                    Else
                        ' skip string not at key position
                        Call Json_ParseString(s, i)
                    End If
                Case "{": depth = depth + 1: i = i + 1
                Case "}": depth = depth - 1: i = i + 1
                Case "[": depth = depth + 1: i = i + 1
                Case "]": depth = depth - 1: i = i + 1
                Case Else: i = i + 1
            End Select
        End If
    Loop
    Json_FindKeyAtTopLevel = 0
End Function

Function JsonValue(ByVal json As String, ByVal key As String) As String
    Dim pos As Long, ch As String, endPos As Long
    If Len(json) = 0 Then Exit Function
    ' If this is an object, ensure we start scanning with depth awareness
    pos = Json_FindKeyAtTopLevel(json, key)
    If pos = 0 Then Exit Function
    ch = Mid$(json, pos, 1)
    Select Case ch
        Case """"  ' string
            JsonValue = Json_ParseString(json, pos)
        Case "["  ' array
            endPos = Json_FindMatching(json, pos, "[", "]")
            If endPos > 0 Then JsonValue = Mid$(json, pos, endPos - pos + 1)
        Case "{"  ' object
            endPos = Json_FindMatching(json, pos, "{", "}")
            If endPos > 0 Then JsonValue = Mid$(json, pos, endPos - pos + 1)
        Case Else ' literal: number, true, false, null
            endPos = pos
            Do While endPos <= Len(json)
                ch = Mid$(json, endPos, 1)
                If ch = "," Or ch = "}" Or ch = "]" Or ch = vbCr Or ch = vbLf Then Exit Do
                endPos = endPos + 1
            Loop
            JsonValue = Trim$(Mid$(json, pos, endPos - pos))
    End Select
End Function

Private Function Json_SplitTopLevelArray(ByVal arrText As String) As Variant
    Dim res() As String
    Dim n As Long, i As Long, ch As String, depth As Long
    Dim inStr As Boolean, esc As Boolean, startEl As Long
    If Len(arrText) < 2 Then
        Json_SplitTopLevelArray = Array()
        Exit Function
    End If
    startEl = 2 ' after [
    For i = 2 To Len(arrText) - 1
        ch = Mid$(arrText, i, 1)
        If inStr Then
            If esc Then
                esc = False
            ElseIf ch = "" Then
                esc = True
            ElseIf ch = """" Then
                inStr = False
            End If
        Else
            Select Case ch
                Case """" : inStr = True
                Case "[", "{" : depth = depth + 1
                Case "]", "}" : depth = depth - 1
                Case ","
                    If depth = 0 Then
                        ReDim Preserve res(n)
                        res(n) = Trim$(Mid$(arrText, startEl, i - startEl))
                        n = n + 1
                        startEl = i + 1
                    End If
            End Select
        End If
    Next i
    If startEl <= Len(arrText) - 1 Then
        ReDim Preserve res(n)
        res(n) = Trim$(Mid$(arrText, startEl, (Len(arrText) - 1) - startEl + 1))
    End If
    If n = 0 And (Len(arrText) <= 2 Or Trim$(Mid$(arrText, 2, Len(arrText) - 2)) = "") Then
        Json_SplitTopLevelArray = Array()
    Else
        Json_SplitTopLevelArray = res
    End If
End Function

Private Function JsonHasKey(ByVal json As String, ByVal key As String) As Boolean
    JsonHasKey = (Json_FindKeyAtTopLevel(json, key) > 0)
End Function

Function ParseJsonArray(ByVal arrText As String) As Variant
    ' Robustly parse a flat array of numbers or strings
    Dim elems As Variant, i As Long
    Dim out() As Variant
    elems = Json_SplitTopLevelArray(arrText)
    If IsEmpty(elems) Then
        ParseJsonArray = Array()
        Exit Function
    End If
    ReDim out(0 To UBound(elems))
    For i = 0 To UBound(elems)
        Dim t As String
        t = Trim$(elems(i))
        If Len(t) >= 2 And Left$(t, 1) = """" And Right$(t, 1) = """" Then
            Dim p As Long: p = 1
            out(i) = Json_ParseString(t, p)
        ElseIf IsNumeric(t) Then
            out(i) = CDbl(t)
        Else
            out(i) = t ' boolean/null or literal retained
        End If
    Next i
    ParseJsonArray = out
End Function

Sub ParseSeries(ByVal seriesText As String, ByRef names() As String, ByRef data() As Variant)
    Dim items As Variant
    Dim i As Long
    items = Json_SplitTopLevelArray(seriesText)
    If IsEmpty(items) Then Exit Sub
    ReDim names(0 To UBound(items))
    ReDim data(0 To UBound(items))
    For i = 0 To UBound(items)
        Dim obj As String
        obj = items(i)
        If Left$(obj, 1) <> "{" Then obj = "{" & obj
        If Right$(obj, 1) <> "}" Then obj = obj & "}"
        names(i) = JsonValue(obj, "name")
        data(i) = ParseJsonArray(JsonValue(obj, "data"))
    Next i
End Sub

Function ParseRowArray(ByVal rowsText As String) As Variant
    Dim rows As Variant, i As Long, out() As Variant
    rows = Json_SplitTopLevelArray(rowsText)
    If IsEmpty(rows) Then
        ParseRowArray = Array()
        Exit Function
    End If
    ReDim out(0 To UBound(rows))
    For i = 0 To UBound(rows)
        Dim rowStr As String
        rowStr = rows(i)
        If Left$(rowStr, 1) <> "[" Then rowStr = "[" & rowStr
        If Right$(rowStr, 1) <> "]" Then rowStr = rowStr & "]"
        out(i) = ParseJsonArray(rowStr)
    Next i
    ParseRowArray = out
End Function

' Create chart at placeholder location (macOS-safe)
Sub CreateChartAtPlaceholder(sld As Slide, placeholder As Shape, chartSpec As String)
    On Error Resume Next
    Dim chartShape As Shape
    Dim chartObj As Object
    Dim l As Single, t As Single, w As Single, h As Single
    Dim chartType As Long
    Dim chartTypeStr As String

    ' Get placeholder dimensions
    l = placeholder.Left
    t = placeholder.Top
    w = placeholder.Width
    h = placeholder.Height

    ' Delete placeholder
    placeholder.Delete

    ' Determine chart type from spec
    chartTypeStr = LCase(JsonValue(chartSpec, "type"))
    Select Case chartTypeStr
        Case "line": chartType = xlLine
        Case "bar": chartType = xlBarClustered
        Case "pie": chartType = xlPie
        Case "area": chartType = xlArea
        Case "scatter": chartType = xlXYScatter
        Case Else
            LogError "E1009", "Unsupported chart type '" & chartTypeStr & "' - defaulting to column"
            chartType = xlColumnClustered
    End Select

    ' Create chart
    Set chartShape = sld.Shapes.AddChart(chartType, l, t, w, h)
    If chartShape Is Nothing Then
        LogError "E1003", "Failed to create chart shape"
        Exit Sub
    End If
    Set chartObj = chartShape.Chart

    Dim xVals As Variant
    Dim seriesNames() As String
    Dim seriesData() As Variant
    Dim dataObj As String
    dataObj = JsonValue(chartSpec, "data")
    If Len(dataObj) = 0 Then
        LogError "E1005", "Chart missing 'data' object"
        Exit Sub
    End If
    If Not JsonHasKey(dataObj, "x") Then
        LogError "E1005", "Chart data missing 'x' categories"
        Exit Sub
    End If
    If Not JsonHasKey(dataObj, "series") Then
        LogError "E1005", "Chart data missing 'series'"
        Exit Sub
    End If
    xVals = ParseJsonArray(JsonValue(dataObj, "x"))
    ParseSeries JsonValue(dataObj, "series"), seriesNames, seriesData

    With chartObj.ChartData
        .Activate
        Dim ws As Object
        Set ws = Nothing
        Set ws = .Workbook.Worksheets(1)
        If ws Is Nothing Then
            LogError "E1011", "Chart data workbook not available (ChartData.Workbook)"
            Exit Sub
        End If
        ws.Cells.Clear

        Dim i As Long, j As Long
        ws.Cells(1, 1).Value = "Category"
        For i = 0 To UBound(seriesNames)
            ws.Cells(1, i + 2).Value = seriesNames(i)
        Next i

        For j = 0 To UBound(xVals)
            ws.Cells(j + 2, 1).Value = xVals(j)
            For i = 0 To UBound(seriesNames)
                ws.Cells(j + 2, i + 2).Value = seriesData(i)(j)
            Next i
        Next j

        On Error Resume Next
        chartObj.SetSourceData Source:=ws.Range(ws.Cells(1, 1), ws.Cells(UBound(xVals) + 2, UBound(seriesNames) + 2))
        If Err.Number <> 0 Then
            LogError "E1011", "Failed to set chart source data: " & Err.Description
            Err.Clear
        End If
        .Workbook.Close
    End With

    On Error GoTo 0
End Sub

' Create table at placeholder location
Sub CreateTableAtPlaceholder(sld As Slide, placeholder As Shape, tableSpec As String)
    On Error Resume Next
    Dim tblShape As Shape
    Dim tbl As Table
    Dim l As Single, t As Single, w As Single, h As Single
    Dim headers As Variant, rows As Variant
    Dim r As Long, c As Long

    ' Get placeholder dimensions
    l = placeholder.Left
    t = placeholder.Top
    w = placeholder.Width
    h = placeholder.Height

    ' Delete placeholder
    placeholder.Delete

    ' Parse table spec
    If Not JsonHasKey(tableSpec, "headers") Then
        LogError "E1005", "Table missing 'headers'"
        Exit Sub
    End If
    If Not JsonHasKey(tableSpec, "rows") Then
        LogError "E1005", "Table missing 'rows'"
        Exit Sub
    End If
    headers = ParseJsonArray(JsonValue(tableSpec, "headers"))
    rows = ParseRowArray(JsonValue(tableSpec, "rows"))

    ' Create table with header row
    Set tblShape = sld.Shapes.AddTable(UBound(rows) + 2, UBound(headers) + 1, l, t, w, h)
    Set tbl = tblShape.Table

    ' Populate headers
    For c = 0 To UBound(headers)
        tbl.Cell(1, c + 1).Shape.TextFrame.TextRange.text = headers(c)
    Next c

    ' Populate rows
    For r = 0 To UBound(rows)
        For c = 0 To UBound(headers)
            tbl.Cell(r + 2, c + 1).Shape.TextFrame.TextRange.text = rows(r)(c)
        Next c
    Next r

    On Error GoTo 0
End Sub



' ================================================================
' MAIN SUBROUTINE
' ================================================================

Sub Main()
    On Error GoTo ErrorHandler
    InitErrorLog

    ' Validate environment
    If Application.Presentations.Count = 0 Then
        LogError "E1007", "No active presentation open"
        ShowErrors
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
    requiredLayouts = Array(1, 2, 4)

    For Each layoutIndex In requiredLayouts
        If Not CacheHas(CLng(layoutIndex)) Then
            Set cl = GetCustomLayoutByIndexSafe(CLng(layoutIndex))
            If cl Is Nothing Then
                LogError "E1008", "Layout index " & layoutIndex & " not found in template"
            Else
                CachePut CLng(layoutIndex), cl
            End If
        End If
    Next layoutIndex

    ' Create slides
    ' ---- Slide 1: Test Presentation ----
    Set currentSlide = AddSlideWithLayout(CacheGet(1))

    ' CenterTitle placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 3, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 1: Missing placeholder Type=CenterTitle (type_id=3), Ordinal=0"
    Else
        SafeSetText shp, "Welcome to the Test"
    End If

    ' Subtitle placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 4, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 1: Missing placeholder Type=Subtitle (type_id=4), Ordinal=0"
    Else
        SafeSetText shp, "A Simple Example"
    End If

    ' ---- Slide 2: Agenda ----
    Set currentSlide = AddSlideWithLayout(CacheGet(2))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 2: Missing placeholder Type=Title (type_id=1), Ordinal=0"
    Else
        SafeSetText shp, "Agenda"
    End If

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 2: Missing placeholder Type=Body (type_id=2), Ordinal=0"
    Else
        SafeSetText shp, "• Introduction" & vbCrLf & "• Main Content" & vbCrLf & "• Conclusion"
    End If

    ' ---- Slide 3: Comparison ----
    Set currentSlide = AddSlideWithLayout(CacheGet(4))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 3: Missing placeholder Type=Title (type_id=1), Ordinal=0"
    Else
        SafeSetText shp, "Comparison"
    End If

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 3: Missing placeholder Type=Body (type_id=2), Ordinal=0"
    Else
        SafeSetText shp, "Left Side:" & vbCrLf & "• Point 1" & vbCrLf & "• Point 2"
    End If

    ' Body placeholder (ordinal 1)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 1)
    If shp Is Nothing Then
        LogError "E1002", "Slide 3: Missing placeholder Type=Body (type_id=2), Ordinal=1"
    Else
        SafeSetText shp, "Right Side:" & vbCrLf & "• Point A" & vbCrLf & "• Point B"
    End If


    ' Report outcome
    If ErrorsCount() > 0 Then
        ShowErrors
    Else
        MsgBox "Successfully created 3 slides!", vbInformation
    End If
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

    requiredLayouts = Array(1, 2, 4)

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