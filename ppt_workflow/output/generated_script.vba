' ================================================================
' AUTO-GENERATED POWERPOINT VBA SCRIPT FOR MAC
' Generated: 2025-09-17 09:59:10 UTC
' Template: ic-template-1.pptx
' Platform: macOS PowerPoint
' ================================================================
'
' USAGE:
'   1. Open your PowerPoint template
'   2. Press Opt+F11 to open VBA editor
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

' Module-level cache for layouts (Collection for Mac)
Dim layoutCache As Collection
Dim errorLog As Collection
Dim warningLog As Collection

' Check if layout cache has a key
Private Function CacheHas(key As Long) As Boolean
    On Error GoTo NotFound
    Dim tmp As Object
    Set tmp = layoutCache(CStr(key))
    CacheHas = True
    Exit Function
NotFound:
    CacheHas = False
End Function

' Add layout to cache
Private Sub CachePut(key As Long, cl As Object)
    On Error Resume Next
    ' Remove if exists
    layoutCache.Remove CStr(key)
    On Error GoTo 0
    layoutCache.Add cl, CStr(key)
End Sub

' Get layout from cache
Private Function CacheGet(key As Long) As Object
    On Error Resume Next
    Set CacheGet = layoutCache(CStr(key))
    On Error GoTo 0
End Function

' Get custom layout by index with fallback mechanism
Function GetCustomLayoutByIndexSafe(layoutIndex As Long) As Object
    On Error Resume Next
    Dim pres As Object
    Dim design As Object
    Dim fallbackLayout As Object

    Set pres = Application.ActivePresentation

    ' First try: Direct index from the active SlideMaster
    If layoutIndex >= 1 And layoutIndex <= pres.SlideMaster.CustomLayouts.Count Then
        Set GetCustomLayoutByIndexSafe = pres.SlideMaster.CustomLayouts(layoutIndex)
        If Not GetCustomLayoutByIndexSafe Is Nothing Then
            Exit Function
        End If
    End If

    ' Second try: Check each Design's SlideMaster
    For Each design In pres.Designs
        If layoutIndex >= 1 And layoutIndex <= design.SlideMaster.CustomLayouts.Count Then
            Set GetCustomLayoutByIndexSafe = design.SlideMaster.CustomLayouts(layoutIndex)
            If Not GetCustomLayoutByIndexSafe Is Nothing Then
                Exit Function
            End If
        End If
    Next design

    ' Fallback: Try to find a similar layout by name pattern
    ' This helps when layout indices shift
    If layoutIndex = 2 Then ' Title Slide
        Set fallbackLayout = FindLayoutByNamePattern("Title")
    ElseIf layoutIndex >= 3 And layoutIndex <= 6 Then ' Content layouts
        Set fallbackLayout = FindLayoutByNamePattern("Content")
    End If

    If Not fallbackLayout Is Nothing Then
        Set GetCustomLayoutByIndexSafe = fallbackLayout
        LogError "W1001", "Layout index " & layoutIndex & " not found, using fallback: " & fallbackLayout.Name
    Else
        Set GetCustomLayoutByIndexSafe = Nothing
    End If

    On Error GoTo 0
End Function

' Find layout by name pattern (fallback mechanism)
Function FindLayoutByNamePattern(pattern As String) As Object
    On Error Resume Next
    Dim pres As Object
    Dim layout As Object

    Set pres = Application.ActivePresentation

    For Each layout In pres.SlideMaster.CustomLayouts
        If InStr(1, layout.Name, pattern, vbTextCompare) > 0 Then
            Set FindLayoutByNamePattern = layout
            Exit Function
        End If
    Next layout

    Set FindLayoutByNamePattern = Nothing
    On Error GoTo 0
End Function

' Add slide with specified layout
Function AddSlideWithLayout(layout As Object) As Object
    Dim pres As Object
    Set pres = Application.ActivePresentation
    Set AddSlideWithLayout = pres.Slides.AddSlide(pres.Slides.Count + 1, layout)
End Function

' Ensure slide is active before operations (important for Mac chart APIs)
Private Sub EnsureSlideActive(sld As Object)
    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then
        Application.ActiveWindow.View.GotoSlide sld.SlideIndex
    ElseIf Application.SlideShowWindows.Count > 0 Then
        Application.SlideShowWindows(1).View.GotoSlide sld.SlideIndex
    End If
    On Error GoTo 0
End Sub

' Get placeholder by type and ordinal (0-based)
Function GetPlaceholderByTypeAndOrdinal(sld As Object, typeId As Long, ordinal As Long) As Object
    On Error Resume Next
    Dim shp As Object
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

' Sort shapes by position using array-based approach (more stable on Mac)
Function SortShapesByPosition(shapes As Collection) As Collection
    If shapes.Count <= 1 Then
        Set SortShapesByPosition = shapes
        Exit Function
    End If

    Dim sorted As Collection
    Dim shpArray() As Object
    Dim i As Long, j As Long
    Dim tempShp As Object

    ' Convert collection to array for more reliable sorting
    ReDim shpArray(1 To shapes.Count)
    For i = 1 To shapes.Count
        Set shpArray(i) = shapes(i)
    Next i

    ' Bubble sort on array
    For i = 1 To UBound(shpArray) - 1
        For j = i + 1 To UBound(shpArray)
            If shpArray(i).Top > shpArray(j).Top Or _
               (Abs(shpArray(i).Top - shpArray(j).Top) < 5 And shpArray(i).Left > shpArray(j).Left) Then
                ' Swap (with 5-pixel tolerance for top alignment)
                Set tempShp = shpArray(i)
                Set shpArray(i) = shpArray(j)
                Set shpArray(j) = tempShp
            End If
        Next j
    Next i

    ' Convert back to collection
    Set sorted = New Collection
    For i = 1 To UBound(shpArray)
        sorted.Add shpArray(i)
    Next i

    Set SortShapesByPosition = sorted
End Function

' Set text with TextFrame2 fallback
Sub SafeSetText(shp As Object, text As String)
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
Sub DebugListPlaceholders(s As Object)
    Dim sh As Object
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
    Set warningLog = New Collection
End Sub

Sub LogError(code As String, details As String)
    On Error Resume Next
    If errorLog Is Nothing Then Set errorLog = New Collection
    errorLog.Add code & ": " & details
    On Error GoTo 0
End Sub

Sub LogWarning(code As String, details As String)
    On Error Resume Next
    If warningLog Is Nothing Then Set warningLog = New Collection
    warningLog.Add code & ": " & details
    On Error GoTo 0
End Sub

Function ErrorsCount() As Long
    If errorLog Is Nothing Then
        ErrorsCount = 0
    Else
        ErrorsCount = errorLog.Count
    End If
End Function

Function WarningsCount() As Long
    If warningLog Is Nothing Then
        WarningsCount = 0
    Else
        WarningsCount = warningLog.Count
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

Sub ShowWarnings()
    If warningLog Is Nothing Or warningLog.Count = 0 Then Exit Sub
    Dim i As Long
    Dim msg As String
    msg = "Completed with " & warningLog.Count & " warning(s):" & vbCrLf & vbCrLf
    For i = 1 To warningLog.Count
        msg = msg & "- " & warningLog(i) & vbCrLf
        If i >= 12 Then
            msg = msg & "... (more omitted)" & vbCrLf
            Exit For
        End If
    Next i
    MsgBox msg, vbInformation, "PowerPoint Script Warnings"
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
                Case """: res = res & """
                Case "\": res = res & "\"
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
        ElseIf ch = "\\" Then
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
    Dim i As Long, depth As Long, ch As String, esc As Boolean, insideString As Boolean
    depth = 0
    For i = startPos To Len(s)
        ch = Mid$(s, i, 1)
        If insideString Then
            If esc Then
                esc = False
            ElseIf ch = "\\" Then
                esc = True
            ElseIf ch = """" Then
                insideString = False
            End If
        Else
            If ch = """" Then
                insideString = True
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
    Dim i As Long
    Dim ch As String
    Dim esc As Boolean
    Dim insideString As Boolean
    Dim depth As Long
    Dim k As String
    Dim c As String

    depth = 0
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If insideString Then
            If esc Then
                esc = False
            ElseIf ch = "\\" Then
                esc = True
            ElseIf ch = """" Then
                insideString = False
            End If
            i = i + 1
        Else
            Select Case ch
                Case """"  ' start string (potential key)
                    If depth = 1 Then
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
    Dim insideString As Boolean, esc As Boolean, startEl As Long
    If Len(arrText) < 2 Then
        Json_SplitTopLevelArray = Array()
        Exit Function
    End If
    startEl = 2 ' after [
    For i = 2 To Len(arrText) - 1
        ch = Mid$(arrText, i, 1)
        If insideString Then
            If esc Then
                esc = False
            ElseIf ch = "\\" Then
                esc = True
            ElseIf ch = """" Then
                insideString = False
            End If
        Else
            Select Case ch
                Case """" : insideString = True
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
        ' Try "values" first (new format), then "data" (old format)
        Dim vals As String
        vals = JsonValue(obj, "values")
        If Len(vals) = 0 Then vals = JsonValue(obj, "data")
        data(i) = ParseJsonArray(vals)
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

' Create chart at placeholder location - Mac optimized
Sub CreateChartAtPlaceholder(sld As Object, placeholder As Object, chartSpec As String)
    On Error Resume Next
    Dim chartShape As Object
    Dim chartObj As Object
    Dim l As Single, t As Single, w As Single, h As Single
    Dim chartType As Long
    Dim chartTypeStr As String
    Dim chartApiHint As String
    Dim preferLegacy As Boolean
    Dim styleCandidates As Variant
    Dim layoutOptions As Variant
    Dim styleIdx As Long
    Dim layoutIdx As Long
    Dim created As Boolean
    Dim fallbackBox As Object
    Dim xVals As Variant
    Dim seriesNames() As String
    Dim seriesData() As Variant
    Dim dataObj As String
    Dim hasCategories As Boolean
    Dim catStr As String

    ' Get placeholder dimensions
    l = placeholder.Left
    t = placeholder.Top
    w = placeholder.Width
    h = placeholder.Height

    ' Determine chart type from spec and chart API hints
    chartTypeStr = LCase(JsonValue(chartSpec, "type"))
    chartApiHint = LCase(JsonValue(chartSpec, "_chart_api"))
    preferLegacy = (chartApiHint = "addchart" Or chartApiHint = "legacy")

    Select Case chartTypeStr
        Case "column", "column_clustered", "clustered_column": chartType = xlColumnClustered
        Case "line": chartType = xlLine
        Case "bar": chartType = xlBarClustered
        Case "pie": chartType = xlPie
        Case "area": chartType = xlArea
        Case "scatter": chartType = xlXYScatter
        Case Else
            LogError "E1009", "Unsupported chart type '" & chartTypeStr & "' - defaulting to column"
            chartType = xlColumnClustered
    End Select

    styleCandidates = Array(-1, 201, 0)
    layoutOptions = Array(False, True)
    created = False

    ' Parse data up front so we can use it for chart or fallback table
    dataObj = JsonValue(chartSpec, "data")
    If Len(dataObj) = 0 Then
        LogError "E1005", "Chart missing 'data' object"
        Exit Sub
    End If
    hasCategories = JsonHasKey(dataObj, "categories") Or JsonHasKey(dataObj, "x")
    If Not hasCategories Then
        LogError "E1005", "Chart data missing 'categories' or 'x'"
        Exit Sub
    End If
    If Not JsonHasKey(dataObj, "series") Then
        LogError "E1005", "Chart data missing 'series'"
        Exit Sub
    End If
    catStr = JsonValue(dataObj, "categories")
    If Len(catStr) = 0 Then catStr = JsonValue(dataObj, "x")
    xVals = ParseJsonArray(catStr)
    ParseSeries JsonValue(dataObj, "series"), seriesNames, seriesData

    ' Ensure the slide is active (Mac quirk for chart creation)
    EnsureSlideActive sld
    DoEvents

    ' Remove the placeholder and rely on shape APIs (InsertChart is unreliable on Mac)
    On Error Resume Next
    placeholder.Delete
    On Error GoTo 0
    On Error Resume Next

    ' Legacy AddChart first if hinted
    If preferLegacy Then
        Err.Clear
        Set chartShape = sld.Shapes.AddChart(chartType, l, t, w, h)
        If Err.Number = 0 And Not chartShape Is Nothing Then created = True
    End If

    ' Try AddChart2 with multiple styles/layout flags
    If Not created Then
        For styleIdx = LBound(styleCandidates) To UBound(styleCandidates)
            For layoutIdx = LBound(layoutOptions) To UBound(layoutOptions)
                Err.Clear
                Set chartShape = sld.Shapes.AddChart2(styleCandidates(styleIdx), chartType, l, t, w, h, layoutOptions(layoutIdx))
                If Err.Number = 0 And Not chartShape Is Nothing Then
                    created = True
                    Exit For
                End If
            Next layoutIdx
            If created Then Exit For
        Next styleIdx
    End If

    ' Try AddChart (modern) if we still do not have a shape
    If Not created Then
        Err.Clear
        Set chartShape = sld.Shapes.AddChart(chartType, l, t, w, h)
        If Err.Number = 0 And Not chartShape Is Nothing Then created = True
    End If

    ' Try AddChart without geometry and reposition manually
    If Not created Then
        Err.Clear
        Set chartShape = sld.Shapes.AddChart(chartType)
        If Err.Number = 0 And Not chartShape Is Nothing Then
            chartShape.Left = l
            chartShape.Top = t
            chartShape.Width = w
            chartShape.Height = h
            created = True
        End If
    End If

    ' Try AddChart2 without geometry and reposition
    If Not created Then
        For styleIdx = LBound(styleCandidates) To UBound(styleCandidates)
            For layoutIdx = LBound(layoutOptions) To UBound(layoutOptions)
                Err.Clear
                Set chartShape = sld.Shapes.AddChart2(styleCandidates(styleIdx), chartType)
                If Err.Number = 0 And Not chartShape Is Nothing Then
                    chartShape.Left = l
                    chartShape.Top = t
                    chartShape.Width = w
                    chartShape.Height = h
                    created = True
                    Exit For
                End If
            Next layoutIdx
            If created Then Exit For
        Next styleIdx
    End If

    If Not created Then
        On Error GoTo 0
        LogWarning "W1003", "Chart placeholder converted to data table after AddChart/AddChart2 failures"
        Set fallbackBox = sld.Shapes.AddTable(UBound(xVals) + 2, UBound(seriesNames) + 2, l, t, w, h)
        With fallbackBox.Table
            Dim headerIdx As Long
            .Cell(1, 1).Shape.TextFrame.TextRange.text = "Category"
            For headerIdx = 0 To UBound(seriesNames)
                .Cell(1, headerIdx + 2).Shape.TextFrame.TextRange.text = seriesNames(headerIdx)
            Next headerIdx
            Dim rowIdx As Long
            For rowIdx = 0 To UBound(xVals)
                .Cell(rowIdx + 2, 1).Shape.TextFrame.TextRange.text = xVals(rowIdx)
                For headerIdx = 0 To UBound(seriesNames)
                    .Cell(rowIdx + 2, headerIdx + 2).Shape.TextFrame.TextRange.text = seriesData(headerIdx)(rowIdx)
                Next headerIdx
            Next rowIdx
        End With
        Exit Sub
    End If

    On Error GoTo 0
    chartShape.Left = l
    chartShape.Top = t
    chartShape.Width = w
    chartShape.Height = h
    Set chartObj = chartShape.Chart

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
Sub CreateTableAtPlaceholder(sld As Object, placeholder As Object, tableSpec As String)
    On Error Resume Next
    Dim tblShape As Object
    Dim tbl As Object
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

    Dim pres As Object
    Dim currentSlide As Object
    Dim shp As Object
    Dim cl As Object

    Set pres = Application.ActivePresentation
    
    ' Initialize layout cache
    Set layoutCache = New Collection

    ' Pre-cache layouts for performance
    Dim layoutIndex As Variant
    Dim requiredLayouts As Variant
    requiredLayouts = Array(2, 6, 8, 10, 48, 54, 56, 57, 58)

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
    ' ---- Slide 1: Digital Transformation Strategy ----
    Set currentSlide = AddSlideWithLayout(CacheGet(58))

    ' CenterTitle placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 3, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 1: Missing placeholder Type=CenterTitle (type_id=3), Ordinal=0"
    Else
        SafeSetText shp, "Digital Transformation"
    End If

    ' Subtitle placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 4, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 1: Missing placeholder Type=Subtitle (type_id=4), Ordinal=0"
    Else
        SafeSetText shp, "Driving Innovation in the Modern Enterprise"
    End If

    ' ---- Slide 2: The Digital Revolution ----
    Set currentSlide = AddSlideWithLayout(CacheGet(2))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 2: Missing placeholder Type=Title (type_id=1), Ordinal=0"
    Else
        SafeSetText shp, "The Digital Revolution"
    End If

    ' ---- Slide 3: What is Digital Transformation? ----
    Set currentSlide = AddSlideWithLayout(CacheGet(56))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 3: Missing placeholder Type=Title (type_id=1), Ordinal=0"
    Else
        SafeSetText shp, "What is Digital Transformation?"
    End If

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 3: Missing placeholder Type=Body (type_id=2), Ordinal=0"
    Else
        SafeSetText shp, "• Integrating digital technology into all business areas" & vbCrLf & "• Fundamentally changing operations and value delivery" & vbCrLf & "• Cultural shift toward experimentation and agility" & vbCrLf & "• Customer-centric approach to innovation"
    End If

    ' ---- Slide 4: Key Technologies Driving Change ----
    Set currentSlide = AddSlideWithLayout(CacheGet(6))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 4: Missing placeholder Type=Title (type_id=1), Ordinal=0"
    Else
        SafeSetText shp, "Key Technologies Driving Change"
    End If

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 4: Missing placeholder Type=Body (type_id=2), Ordinal=0"
    Else
        SafeSetText shp, "Cloud Computing:" & vbCrLf & "• Scalable infrastructure" & vbCrLf & "• Cost optimization" & vbCrLf & "• Global accessibility" & vbCrLf & "" & vbCrLf & "Artificial Intelligence:" & vbCrLf & "• Automation" & vbCrLf & "• Predictive analytics" & vbCrLf & "• Enhanced decision-making"
    End If

    ' Body placeholder (ordinal 1)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 1)
    If shp Is Nothing Then
        LogError "E1002", "Slide 4: Missing placeholder Type=Body (type_id=2), Ordinal=1"
    Else
        SafeSetText shp, "Internet of Things:" & vbCrLf & "• Connected devices" & vbCrLf & "• Real-time data" & vbCrLf & "• Smart operations" & vbCrLf & "" & vbCrLf & "Blockchain:" & vbCrLf & "• Security" & vbCrLf & "• Transparency" & vbCrLf & "• Decentralization"
    End If

    ' ---- Slide 5: Digital Adoption Trends ----
    Set currentSlide = AddSlideWithLayout(CacheGet(57))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 5: Missing placeholder Type=Title (type_id=1), Ordinal=0"
    Else
        SafeSetText shp, "Digital Adoption Trends"
    End If

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 5: Missing placeholder Type=Body (type_id=2), Ordinal=0"
    Else
        SafeSetText shp, "• 70% of companies have a digital strategy" & vbCrLf & "• Cloud adoption grew 25% year-over-year" & vbCrLf & "• AI implementation increased by 37%" & vbCrLf & "• Remote work tools usage up 300%"
    End If

    ' Chart placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 8, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 5: Missing placeholder Type=Chart (type_id=8), Ordinal=0"
    Else
        CreateChartAtPlaceholder currentSlide, shp, "{""type"":""column"",""data"":{""categories"":[""2020"",""2021"",""2022"",""2023"",""2024""],""series"":[{""name"":""Cloud"",""values"":[45,56,68,78,85]},{""name"":""AI/ML"",""values"":[20,28,39,48,57]},{""name"":""IoT"",""values"":[15,22,31,38,45]}]},""_chart_api"":""addchart""}"
    End If

    ' ---- Slide 6: Strategic Pillars ----
    Set currentSlide = AddSlideWithLayout(CacheGet(8))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 6: Missing placeholder Type=Title (type_id=1), Ordinal=0"
    Else
        SafeSetText shp, "Strategic Pillars"
    End If

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 6: Missing placeholder Type=Body (type_id=2), Ordinal=0"
    Else
        SafeSetText shp, "Customer Experience" & vbCrLf & "" & vbCrLf & "• Personalization" & vbCrLf & "• Omnichannel presence" & vbCrLf & "• Real-time support" & vbCrLf & "• Data-driven insights"
    End If

    ' Body placeholder (ordinal 1)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 1)
    If shp Is Nothing Then
        LogError "E1002", "Slide 6: Missing placeholder Type=Body (type_id=2), Ordinal=1"
    Else
        SafeSetText shp, "Operational Excellence" & vbCrLf & "" & vbCrLf & "• Process automation" & vbCrLf & "• Efficiency gains" & vbCrLf & "• Cost reduction" & vbCrLf & "• Quality improvement"
    End If

    ' Body placeholder (ordinal 2)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 2)
    If shp Is Nothing Then
        LogError "E1002", "Slide 6: Missing placeholder Type=Body (type_id=2), Ordinal=2"
    Else
        SafeSetText shp, "Business Model Innovation" & vbCrLf & "" & vbCrLf & "• New revenue streams" & vbCrLf & "• Platform economies" & vbCrLf & "• Subscription models" & vbCrLf & "• Digital products"
    End If

    ' ---- Slide 7: Implementation Roadmap ----
    Set currentSlide = AddSlideWithLayout(CacheGet(56))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 7: Missing placeholder Type=Title (type_id=1), Ordinal=0"
    Else
        SafeSetText shp, "Implementation Roadmap"
    End If

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 7: Missing placeholder Type=Body (type_id=2), Ordinal=0"
    Else
        SafeSetText shp, "Phase 1: Assessment (Months 1-2)" & vbCrLf & "• Current state analysis" & vbCrLf & "• Gap identification" & vbCrLf & "• Stakeholder alignment" & vbCrLf & "" & vbCrLf & "Phase 2: Planning (Months 3-4)" & vbCrLf & "• Strategy development" & vbCrLf & "• Resource allocation" & vbCrLf & "• Timeline creation" & vbCrLf & "" & vbCrLf & "Phase 3: Execution (Months 5-12)" & vbCrLf & "• Pilot programs" & vbCrLf & "• Scaling initiatives" & vbCrLf & "• Continuous monitoring"
    End If

    ' ---- Slide 8: Success Factors ----
    Set currentSlide = AddSlideWithLayout(CacheGet(10))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 8: Missing placeholder Type=Title (type_id=1), Ordinal=0"
    Else
        SafeSetText shp, "Success Factors"
    End If

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 8: Missing placeholder Type=Body (type_id=2), Ordinal=0"
    Else
        SafeSetText shp, "Leadership" & vbCrLf & "• Executive sponsorship" & vbCrLf & "• Clear vision" & vbCrLf & "• Change champions"
    End If

    ' Body placeholder (ordinal 1)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 1)
    If shp Is Nothing Then
        LogError "E1002", "Slide 8: Missing placeholder Type=Body (type_id=2), Ordinal=1"
    Else
        SafeSetText shp, "Culture" & vbCrLf & "• Growth mindset" & vbCrLf & "• Innovation focus" & vbCrLf & "• Risk tolerance"
    End If

    ' Body placeholder (ordinal 2)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 2)
    If shp Is Nothing Then
        LogError "E1002", "Slide 8: Missing placeholder Type=Body (type_id=2), Ordinal=2"
    Else
        SafeSetText shp, "Technology" & vbCrLf & "• Modern infrastructure" & vbCrLf & "• Integration capabilities" & vbCrLf & "• Security focus"
    End If

    ' Body placeholder (ordinal 3)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 3)
    If shp Is Nothing Then
        LogError "E1002", "Slide 8: Missing placeholder Type=Body (type_id=2), Ordinal=3"
    Else
        SafeSetText shp, "Talent" & vbCrLf & "• Digital skills" & vbCrLf & "• Continuous learning" & vbCrLf & "• Cross-functional teams"
    End If

    ' ---- Slide 9: Key Takeaway ----
    Set currentSlide = AddSlideWithLayout(CacheGet(48))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 9: Missing placeholder Type=Title (type_id=1), Ordinal=0"
    Else
        SafeSetText shp, "Key Takeaway"
    End If

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 9: Missing placeholder Type=Body (type_id=2), Ordinal=0"
    Else
        SafeSetText shp, """Digital transformation is not about technology—it's about talent, process, and organizational change management.""" & vbCrLf & "" & vbCrLf & "— Thomas Davenport, Analytics Expert"
    End If

    ' ---- Slide 10: Thank You ----
    Set currentSlide = AddSlideWithLayout(CacheGet(54))

    ' Title placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 1, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 10: Missing placeholder Type=Title (type_id=1), Ordinal=0"
    Else
        SafeSetText shp, "Thank You"
    End If

    ' Body placeholder (ordinal 0)
    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)
    If shp Is Nothing Then
        LogError "E1002", "Slide 10: Missing placeholder Type=Body (type_id=2), Ordinal=0"
    Else
        SafeSetText shp, "Questions & Discussion" & vbCrLf & "" & vbCrLf & "Contact:" & vbCrLf & "email@example.com" & vbCrLf & "www.company.com" & vbCrLf & "" & vbCrLf & "Let's Transform Together!"
    End If


    ' Report outcome
    If ErrorsCount() > 0 Then
        ShowErrors
    ElseIf WarningsCount() > 0 Then
        ShowWarnings
    Else
        MsgBox "Successfully created 10 slides!", vbInformation
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
    Dim pres As Object
    Dim layout As Object
    Dim msg As String

    Set pres = Application.ActivePresentation

    msg = "Template Validation Report:" & vbCrLf & vbCrLf
    msg = msg & "Template: " & pres.Name & vbCrLf
    msg = msg & "Layouts: " & pres.SlideMaster.CustomLayouts.Count & vbCrLf
    msg = msg & "Platform: macOS" & vbCrLf & vbCrLf

    ' Check required layouts
    Dim requiredLayouts As Variant
    Dim layoutIndex As Variant
    Dim found As Boolean

    requiredLayouts = Array(2, 6, 8, 10, 48, 54, 56, 57, 58)

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