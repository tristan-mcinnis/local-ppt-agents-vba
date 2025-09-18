"""
Step 2: Convert slide_plan.json to PowerPoint VBA script
Generates macOS-safe VBA code with complete helper functions
"""

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional


class PlanToVBAConverter:
    """Converts slide plan to executable VBA script"""

    def __init__(self, plan_path: str, debug_slide: Optional[int] = None):
        """Initialize with path to slide plan

        Args:
            plan_path: Path to the slide plan JSON file.
            debug_slide: Optional slide number to output placeholder
                diagnostics for. By default, no debugging information is
                generated.
        """
        self.plan = self._load_json(plan_path)
        self.code_lines = []
        self.used_layouts = set()
        self.debug_slide = debug_slide

    @staticmethod
    def _load_json(path: str) -> Dict:
        """Load and parse JSON file"""
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)

    @staticmethod
    def _vba_escape(s: str) -> str:
        """Escape string for VBA"""
        if not s:
            return ""
        # Escape quotes for VBA
        s = s.replace('"', '""')
        # Handle newlines
        s = s.replace('\n', '" & vbCrLf & "')
        return s

    def _generate_header(self) -> str:
        """Generate VBA file header with constants and declarations"""
        now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        template_name = self.plan["meta"]["template_name"]

        return f"""' ================================================================
' AUTO-GENERATED POWERPOINT VBA SCRIPT FOR MAC
' Generated: {now}
' Template: {template_name}
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
"""

    def _generate_helper_functions(self) -> str:
        """Generate all helper functions"""
        return '''
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
                Case "\"": res = res & "\""
                Case "\\": res = res & "\\"
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
        ElseIf ch = "\\\\" Then
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
            ElseIf ch = "\\\\" Then
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
            ElseIf ch = "\\\\" Then
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
            ElseIf ch = "\\\\" Then
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

    ' Try InsertChart on placeholder first (works better on some Mac versions)
    On Error Resume Next
    If placeholder.PlaceholderFormat.Type = ppPlaceholderChart Or _
       placeholder.PlaceholderFormat.Type = ppPlaceholderObject Then
        Err.Clear
        Set chartShape = placeholder.InsertChart2(-1, chartType)
        If Err.Number = 0 And Not chartShape Is Nothing Then
            created = True
        Else
            ' Try legacy InsertChart
            Err.Clear
            Set chartShape = placeholder.InsertChart(chartType)
            If Err.Number = 0 And Not chartShape Is Nothing Then created = True
        End If
    End If

    ' If InsertChart failed, delete placeholder and try AddChart
    If Not created Then
        placeholder.Delete
    End If
    On Error GoTo 0
    On Error Resume Next

    ' Legacy AddChart if hinted or InsertChart failed
    If Not created And preferLegacy Then
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

    ' Try to populate chart data (may fail on Mac)
    On Error Resume Next
    Dim dataActivated As Boolean
    dataActivated = False

    ' Try to activate chart data
    chartObj.ChartData.Activate
    If Err.Number = 0 Then
        dataActivated = True
    Else
        ' Try alternative activation method
        Err.Clear
        chartObj.Activate
        DoEvents
        chartObj.ChartData.Activate
        If Err.Number = 0 Then dataActivated = True
    End If

    If dataActivated Then
        With chartObj.ChartData
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
    Else
        ' Chart data couldn't be activated - log warning but keep the chart
        LogWarning "W1004", "Chart created but data couldn't be populated (Mac limitation)"
    End If

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

'''

    def _generate_slide_code(self, slide: Dict) -> List[str]:
        """Generate VBA code for a single slide"""
        code = []
        slide_num = slide["slide_number"]
        layout_idx = slide["selected_layout"]["index"]

        # Track used layout
        self.used_layouts.add(layout_idx)

        code.append(f"    ' ---- Slide {slide_num}: {slide.get('slide_title', '')} ----")
        code.append(f"    Set currentSlide = AddSlideWithLayout(CacheGet({layout_idx}))")

        # Optional debug output for diagnosing placeholder issues
        if self.debug_slide is not None and slide_num == self.debug_slide:
            code.append("    ")
            code.append(
                f"    ' Debug: List placeholders on slide {slide_num} (layout {layout_idx})"
            )
            code.append("    DebugListPlaceholders currentSlide")

        code.append("")

        # Process each content item
        for content in slide["content_map"]:
            ph_type = content["placeholder_type"]
            type_id = content["type_id"]
            ordinal = content["ordinal"]
            content_type = content["content_type"]
            content_data = content["content_data"]

            # Skip image placeholders entirely
            if content_type == "image_path":
                image_path = self._vba_escape(content_data)
                code.append(f"    ")
                code.append(f"    ' Image placeholder skipped: {image_path}")
                code.append(f"    ' User will add image manually after slides are created")
                code.append("")
                continue

            # Get placeholder - STRICT MATCH (for non-images)
            code.append(f"    ' {ph_type} placeholder (ordinal {ordinal})")
            code.append(f"    Set shp = GetPlaceholderByTypeAndOrdinal(currentSlide, {type_id}, {ordinal})")
            code.append(f"    If shp Is Nothing Then")
            code.append(
                f'        LogError "E1002", "Slide {slide_num}: Missing placeholder Type={ph_type} (type_id={type_id}), Ordinal={ordinal}"')
            code.append(f"    Else")

            if content_type == "text":
                # Handle text content
                text = self._vba_escape(content_data)
                if '\n' in content_data or '•' in content_data or '-' in content_data:
                    # Multi-line or bullet text
                    code.append(f'        SafeSetText shp, "{text}"')
                else:
                    # Simple text
                    code.append(f'        SafeSetText shp, "{text}"')

            elif content_type == "chart":
                # Handle chart - emit compact JSON (no spaces) for simpler VBA parsing
                chart_payload = dict(content_data)
                chart_api_hint = slide.get("platform_hints", {}).get("chart_api")
                if chart_api_hint:
                    chart_payload["_chart_api"] = chart_api_hint.lower()
                chart_json = json.dumps(chart_payload, ensure_ascii=False, separators=(",", ":"))
                chart_escaped = self._vba_escape(chart_json)
                code.append(f'        CreateChartAtPlaceholder currentSlide, shp, "{chart_escaped}"')

            elif content_type == "table":
                # Handle table - emit compact JSON (no spaces) for simpler VBA parsing
                table_json = json.dumps(content_data, ensure_ascii=False, separators=(",", ":"))
                table_escaped = self._vba_escape(table_json)
                code.append(f'        CreateTableAtPlaceholder currentSlide, shp, "{table_escaped}"')
            code.append(f"    End If")
            code.append("")

        return code

    def _generate_main_sub(self) -> str:
        """Generate the main subroutine"""
        # Pre-scan slides to collect layout indices
        for slide in self.plan["slides"]:
            if "selected_layout" in slide and "index" in slide["selected_layout"]:
                self.used_layouts.add(slide["selected_layout"]["index"])

        code = ["", "' ================================================================"]
        code.append("' MAIN SUBROUTINE")
        code.append("' ================================================================")
        code.append("")
        code.append("Sub Main()")
        code.append("    On Error GoTo ErrorHandler")
        code.append("    InitErrorLog")
        code.append("")
        code.append("    ' Validate environment")
        code.append("    If Application.Presentations.Count = 0 Then")
        code.append('        LogError "E1007", "No active presentation open"')
        code.append("        ShowErrors")
        code.append("        Exit Sub")
        code.append("    End If")
        code.append("")
        code.append("    Dim pres As Object")
        code.append("    Dim currentSlide As Object")
        code.append("    Dim shp As Object")
        code.append("    Dim cl As Object")
        code.append("")
        code.append("    Set pres = Application.ActivePresentation")
        code.append("    ")
        code.append("    ' Initialize layout cache")
        code.append("    Set layoutCache = New Collection")
        code.append("")
        code.append("    ' Pre-cache layouts for performance")
        code.append("    Dim layoutIndex As Variant")
        code.append("    Dim requiredLayouts As Variant")

        # Add layout indices - now self.used_layouts is populated
        layout_indices = sorted(list(self.used_layouts))
        code.append(f"    requiredLayouts = Array({', '.join(map(str, layout_indices))})")
        code.append("")
        code.append("    For Each layoutIndex In requiredLayouts")
        code.append("        If Not CacheHas(CLng(layoutIndex)) Then")
        code.append("            Set cl = GetCustomLayoutByIndexSafe(CLng(layoutIndex))")
        code.append("            If cl Is Nothing Then")
        code.append('                LogError "E1008", "Layout index " & layoutIndex & " not found in template"')
        code.append("            Else")
        code.append("                CachePut CLng(layoutIndex), cl")
        code.append("            End If")
        code.append("        End If")
        code.append("    Next layoutIndex")
        code.append("")
        code.append("    ' Create slides")

        # Generate code for each slide
        for slide in self.plan["slides"]:
            slide_code = self._generate_slide_code(slide)
            code.extend(slide_code)

        code.append("")
        code.append("    ' Report outcome")
        code.append("    If ErrorsCount() > 0 Then")
        code.append("        ShowErrors")
        code.append("    ElseIf WarningsCount() > 0 Then")
        code.append("        ShowWarnings")
        code.append("    Else")
        code.append(f'        MsgBox "Successfully created {len(self.plan["slides"])} slides!", vbInformation')
        code.append("    End If")
        code.append("    Exit Sub")
        code.append("")
        code.append("ErrorHandler:")
        code.append('    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical')
        code.append("End Sub")

        return "\n".join(code)

    def _generate_validation_sub(self) -> str:
        """Generate optional validation subroutine"""
        return '''

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

    requiredLayouts = Array(''' + ', '.join(map(str, sorted(self.used_layouts))) + ''')

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
End Sub'''

    def convert(self) -> str:
        """Convert plan to VBA script"""
        # Generate all sections
        sections = [
            self._generate_header(),
            self._generate_helper_functions(),
            self._generate_main_sub(),
            self._generate_validation_sub()
        ]

        return "\n".join(sections)


def main():
    """CLI entry point"""
    import sys

    if len(sys.argv) != 3:
        print("Usage: python plan_to_vba.py slide_plan.json output_script.vba")
        sys.exit(1)

    plan_path = sys.argv[1]
    output_path = sys.argv[2]

    # Convert
    converter = PlanToVBAConverter(plan_path)
    vba_code = converter.convert()

    # Write output
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(vba_code)

    # Report results
    slide_count = len(converter.plan["slides"])
    layout_count = len(converter.used_layouts)
    print(f"✓ Generated VBA script for {slide_count} slides using {layout_count} layouts")
    print(f"✓ Script saved to: {output_path}")
    print("\nNext steps:")
    print("1. Open your PowerPoint template")
    print("2. Press Opt+F11 (Mac)")
    print("3. Insert > Module")
    print("4. Paste the generated script")
    print("5. Run 'ValidateTemplate' to check compatibility")
    print("6. Run 'Main' to create slides")


if __name__ == "__main__":
    main()
