"""
Step 2: Convert slide_plan.json to PowerPoint VBA script
Generates macOS-safe VBA code with complete helper functions
"""

import json
from datetime import datetime
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
        now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
        template_name = self.plan["meta"]["template_name"]

        return f"""' ================================================================
' AUTO-GENERATED POWERPOINT VBA SCRIPT
' Generated: {now}
' Template: {template_name}
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
"""

    def _generate_helper_functions(self) -> str:
        """Generate all helper functions"""
        return '''
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
            code.append(f'        MsgBox "STRICT MATCH ERROR: Missing required placeholder on slide {slide_num}:" & vbCrLf & _')
            code.append(f'               "Type: {ph_type} (type_id={type_id})" & vbCrLf & _')
            code.append(f'               "Ordinal: {ordinal}" & vbCrLf & vbCrLf & _')
            code.append(f'               "This placeholder is required but not found in the layout.", vbCritical, "Missing Placeholder"')
            code.append(f"        Exit Sub")
            code.append(f"    End If")

            if content_type == "text":
                # Handle text content
                text = self._vba_escape(content_data)
                if '\n' in content_data or '•' in content_data or '-' in content_data:
                    # Multi-line or bullet text
                    code.append(f'    SafeSetText shp, "{text}"')
                else:
                    # Simple text
                    code.append(f'    SafeSetText shp, "{text}"')

            elif content_type == "chart":
                # Handle chart
                chart_json = json.dumps(content_data, ensure_ascii=False)
                chart_escaped = self._vba_escape(chart_json)
                code.append(f'    CreateChartAtPlaceholder currentSlide, shp, "{chart_escaped}"')

            elif content_type == "table":
                # Handle table
                table_json = json.dumps(content_data, ensure_ascii=False)
                table_escaped = self._vba_escape(table_json)
                code.append(f'    CreateTableAtPlaceholder currentSlide, shp, "{table_escaped}"')
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
        code.append("")
        code.append("    ' Validate environment")
        code.append("    If Application.Presentations.Count = 0 Then")
        code.append('        MsgBox "Please open a PowerPoint presentation first.", vbExclamation')
        code.append("        Exit Sub")
        code.append("    End If")
        code.append("")
        code.append("    Dim pres As Presentation")
        code.append("    Dim currentSlide As Slide")
        code.append("    Dim shp As Shape")
        code.append("    Dim cl As CustomLayout")
        code.append("")
        code.append("    Set pres = Application.ActivePresentation")
        code.append("    ")
        code.append("    ' Initialize layout cache (macOS-safe Collection)")
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
        code.append('                MsgBox "Layout index " & layoutIndex & " not found in template. Check that you have the correct template open.", vbCritical')
        code.append("                Exit Sub")
        code.append("            End If")
        code.append("            CachePut CLng(layoutIndex), cl")
        code.append("        End If")
        code.append("    Next layoutIndex")
        code.append("")
        code.append("    ' Create slides")

        # Generate code for each slide
        for slide in self.plan["slides"]:
            slide_code = self._generate_slide_code(slide)
            code.extend(slide_code)

        code.append("")
        code.append("    ' Success message")
        code.append(f'    MsgBox "Successfully created {len(self.plan["slides"])} slides!", vbInformation')
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
    print("2. Press Alt+F11 (Windows) or Opt+F11 (Mac)")
    print("3. Insert > Module")
    print("4. Paste the generated script")
    print("5. Run 'ValidateTemplate' to check compatibility")
    print("6. Run 'Main' to create slides")


if __name__ == "__main__":
    main()