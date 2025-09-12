' ================================================================
' UNIVERSAL POWERPOINT TEMPLATE ANALYZER (Enhanced Version)
' Version 3.1 - Fixed Error 438 compatibility issues
' Generates comprehensive, token-optimized JSON analysis
' Fully cross-platform compatible (Windows & macOS)
' Enhanced with geometry data and validation
' ================================================================

Option Explicit

' --- Platform detection and constants
#If Mac Then
    Const PATH_SEP As String = "/"
    Const PLATFORM As String = "macOS"
#Else
    Const PATH_SEP As String = "\"
    Const PLATFORM As String = "Windows"
#End If

' PowerPoint constants for better compatibility
Const msoPlaceholder = 14
Const ppPlaceholderTitle = 1
Const ppPlaceholderBody = 2
Const ppPlaceholderCenterTitle = 3
Const ppPlaceholderSubtitle = 4
Const ppPlaceholderObject = 7
Const ppPlaceholderChart = 8
Const ppPlaceholderTable = 9
Const ppPlaceholderClipArt = 10
Const ppPlaceholderMedia = 12
Const ppPlaceholderSlideImage = 13
Const ppPlaceholderPicture = 18
Const ppPlaceholderContent = 19

' --- Helper function to escape strings for JSON
Private Function JsonEscape(str As String) As String
    If IsNull(str) Or str = "" Then
        JsonEscape = ""
        Exit Function
    End If
    
    str = Replace(str, "\", "\\")
    str = Replace(str, """", "\""")
    str = Replace(str, vbCr, "\r")
    str = Replace(str, vbLf, "\n")
    str = Replace(str, vbTab, "\t")
    str = Replace(str, "/", "\/")
    JsonEscape = str
End Function

' --- Cross-platform helper to get a filename without its extension
Private Function GetBaseName(fileName As String) As String
    Dim dotPos As Integer
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetBaseName = Left(fileName, dotPos - 1)
    Else
        GetBaseName = fileName
    End If
End Function

' --- Format a number to 2 decimal places for JSON
Private Function FormatNumber2(num As Single) As String
    FormatNumber2 = Format(num, "0.00")
End Function

' --- Safely get theme name with error handling
Private Function GetThemeName(ppt As Presentation) As String
    On Error Resume Next
    Dim themeName As String
    themeName = ""
    
    ' Try different methods to get theme name
    ' Method 1: Try direct theme access (may not work in all versions)
    themeName = ppt.SlideMaster.Theme.Name
    
    ' If that failed, try alternative
    If themeName = "" Or Err.Number <> 0 Then
        Err.Clear
        ' Method 2: Try through Design property (older PowerPoint versions)
        themeName = ppt.Designs(1).Name
    End If
    
    ' If still failed, use default
    If themeName = "" Or Err.Number <> 0 Then
        themeName = "Default Theme"
    End If
    
    On Error GoTo 0
    GetThemeName = themeName
End Function

' --- Main Analysis Subroutine
Sub UniversalTemplateAnalyzer()
    Dim ppt As Presentation
    Dim outputPath As String
    Dim jsonContent As String
    Dim layoutsArray As Collection
    Dim validationNotes As Collection
    
    Set layoutsArray = New Collection
    Set validationNotes = New Collection
    
    On Error GoTo ErrorHandler
    
    ' Validate that a presentation is open
    If Application.Presentations.Count = 0 Then
        MsgBox "No presentation is open. Please open a PowerPoint file first.", vbExclamation
        Exit Sub
    End If
    
    Set ppt = ActivePresentation
    
    ' Validate presentation has content
    If ppt.SlideMaster.CustomLayouts.Count = 0 Then
        MsgBox "This presentation has no custom layouts to analyze.", vbExclamation
        Exit Sub
    End If
    
    ' Define output path for the JSON file
    outputPath = ppt.Path & PATH_SEP & GetBaseName(ppt.Name) & "_analysis.json"
    
    ' --- Build Enhanced Template Info JSON Object
    Dim templateInfo As String
    Dim themeName As String
    themeName = GetThemeName(ppt)
    
    templateInfo = """template_info"":{" & _
                   """name"":""" & JsonEscape(ppt.Name) & """," & _
                   """path"":""" & JsonEscape(ppt.FullName) & """," & _
                   """analysis_date"":""" & Format(Now(), "yyyy-mm-dd") & "T" & Format(Now(), "hh:mm:ss") & "Z""," & _
                   """analyzer_version"":""3.1""," & _
                   """platform"":""" & PLATFORM & """," & _
                   """slide_count"":" & ppt.Slides.Count & "," & _
                   """slide_master"":{" & _
                   """name"":""" & JsonEscape(ppt.SlideMaster.Name) & """," & _
                   """layout_count"":" & ppt.SlideMaster.CustomLayouts.Count & "," & _
                   """theme_name"":""" & JsonEscape(themeName) & """" & _
                   "}}"
                   
    ' --- Process each layout
    Dim layout As CustomLayout
    Dim layoutIndex As Long
    layoutIndex = 0
    
    For Each layout In ppt.SlideMaster.CustomLayouts
        layoutIndex = layoutIndex + 1
        Dim layoutPlaceholders As Collection
        Set layoutPlaceholders = New Collection
        
        Dim shp As Shape
        Dim placeholderCount As Long
        placeholderCount = 0
        
        ' Count and process placeholders
        For Each shp In layout.Shapes
            If shp.Type = msoPlaceholder Then
                placeholderCount = placeholderCount + 1
                
                ' Build enhanced placeholder JSON with geometry
                Dim phInfo As String
                phInfo = BuildPlaceholderJSON(shp, placeholderCount - 1)
                layoutPlaceholders.Add phInfo
            End If
        Next shp
        
        ' --- Join placeholders into an array string
        Dim placeholdersJson As String
        placeholdersJson = ""
        If layoutPlaceholders.Count > 0 Then
            placeholdersJson = JoinCollection(layoutPlaceholders, ",")
        End If
        
        ' --- Determine layout category
        Dim layoutCategory As String
        layoutCategory = CategorizeLayout(layout.Name, placeholderCount)
        
        ' --- Build enhanced JSON object for this layout
        Dim layoutJson As String
        layoutJson = "{" & _
                     """index"":" & layout.Index & "," & _
                     """name"":""" & JsonEscape(layout.Name) & """," & _
                     """category"":""" & layoutCategory & """," & _
                     """placeholder_count"":" & placeholderCount & "," & _
                     """is_blank"":" & IIf(placeholderCount = 0, "true", "false") & "," & _
                     """placeholders"":[" & placeholdersJson & "]" & _
                     "}"
                     
        layoutsArray.Add layoutJson
        
        ' Add validation note if layout has issues
        If placeholderCount = 0 And InStr(LCase(layout.Name), "blank") = 0 Then
            validationNotes.Add """Layout '" & JsonEscape(layout.Name) & "' has no placeholders but is not marked as blank"""
        End If
    Next layout
    
    ' --- Build validation notes array
    Dim validationJson As String
    If validationNotes.Count > 0 Then
        validationJson = ",""validation_notes"":[" & JoinCollection(validationNotes, ",") & "]"
    Else
        validationJson = ",""validation_notes"":[]"
    End If
    
    ' --- Add analysis statistics
    Dim statsJson As String
    statsJson = ",""statistics"":{" & _
                """total_layouts"":" & layoutsArray.Count & "," & _
                """layouts_with_placeholders"":" & CountLayoutsWithPlaceholders(ppt) & "," & _
                """average_placeholders_per_layout"":" & FormatNumber2(GetAveragePlaceholders(ppt)) & _
                "}"
    
    ' --- Assemble the final enhanced JSON string
    jsonContent = "{" & _
                  templateInfo & "," & _
                  """layouts"":[" & JoinCollection(layoutsArray, ",") & "]" & _
                  statsJson & _
                  validationJson & _
                  "}"
                  
    ' --- Write JSON content to file
    Dim fileNum As Integer
    fileNum = FreeFile
    Open outputPath For Output As #fileNum
    Print #fileNum, jsonContent
    Close #fileNum
    
    ' --- Success Message with details
    MsgBox "Template Analysis Complete!" & vbCrLf & vbCrLf & _
           "Version: 3.1 (Enhanced)" & vbCrLf & _
           "Layouts analyzed: " & layoutsArray.Count & vbCrLf & _
           "Output saved to:" & vbCrLf & outputPath & vbCrLf & vbCrLf & _
           "This analysis includes geometry data and validation notes.", _
           vbInformation, "Analysis Successful"

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during analysis:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Location: " & Err.Source & vbCrLf & vbCrLf & _
           "Please ensure you have write permissions to:" & vbCrLf & _
           ppt.Path, vbCritical, "Analysis Error"
End Sub

' --- Build placeholder JSON with safe property access
Private Function BuildPlaceholderJSON(shp As Shape, idx As Long) As String
    On Error Resume Next
    Dim result As String
    Dim placeholderType As Long
    
    ' Safely get placeholder type
    placeholderType = 0
    placeholderType = shp.PlaceholderFormat.Type
    If Err.Number <> 0 Then
        placeholderType = 0
        Err.Clear
    End If
    
    ' Build JSON string with safe property access
    result = "{"
    result = result & """id"":" & shp.Id & ","
    result = result & """type_name"":""" & GetPlaceholderTypeName(placeholderType) & ""","
    result = result & """type_id"":" & placeholderType & ","
    result = result & """index"":" & idx & ","
    result = result & """geometry"":{"
    result = result & """left"":" & FormatNumber2(shp.Left) & ","
    result = result & """top"":" & FormatNumber2(shp.Top) & ","
    result = result & """width"":" & FormatNumber2(shp.Width) & ","
    result = result & """height"":" & FormatNumber2(shp.Height)
    result = result & "},"
    result = result & """has_text_frame"":" & IIf(HasTextFrame(shp), "true", "false")
    result = result & "}"
    
    On Error GoTo 0
    BuildPlaceholderJSON = result
End Function

' --- Check if a shape has a text frame
Private Function HasTextFrame(shp As Shape) As Boolean
    On Error Resume Next
    Dim test As Object
    Dim hasFrame As Boolean
    hasFrame = False
    
    ' Try to access TextFrame
    Set test = shp.TextFrame
    If Err.Number = 0 And Not test Is Nothing Then
        hasFrame = True
    End If
    
    Err.Clear
    On Error GoTo 0
    HasTextFrame = hasFrame
End Function

' --- Categorize layout based on name and structure
Private Function CategorizeLayout(layoutName As String, placeholderCount As Long) As String
    Dim lName As String
    lName = LCase(layoutName)
    
    If InStr(lName, "title") > 0 And InStr(lName, "slide") > 0 Then
        CategorizeLayout = "title"
    ElseIf InStr(lName, "blank") > 0 Or placeholderCount = 0 Then
        CategorizeLayout = "blank"
    ElseIf InStr(lName, "two") > 0 Or InStr(lName, "2") > 0 Then
        CategorizeLayout = "two-content"
    ElseIf InStr(lName, "comparison") > 0 Then
        CategorizeLayout = "comparison"
    ElseIf InStr(lName, "section") > 0 Then
        CategorizeLayout = "section"
    ElseIf InStr(lName, "picture") > 0 Or InStr(lName, "image") > 0 Then
        CategorizeLayout = "picture"
    ElseIf InStr(lName, "chart") > 0 Then
        CategorizeLayout = "chart"
    ElseIf InStr(lName, "table") > 0 Then
        CategorizeLayout = "table"
    Else
        CategorizeLayout = "content"
    End If
End Function

' --- Count layouts that have placeholders
Private Function CountLayoutsWithPlaceholders(ppt As Presentation) As Long
    On Error Resume Next
    Dim count As Long
    Dim layout As CustomLayout
    Dim shp As Shape
    
    count = 0
    For Each layout In ppt.SlideMaster.CustomLayouts
        For Each shp In layout.Shapes
            If shp.Type = msoPlaceholder Then
                count = count + 1
                Exit For ' Move to next layout once we find a placeholder
            End If
        Next shp
    Next layout
    
    On Error GoTo 0
    CountLayoutsWithPlaceholders = count
End Function

' --- Calculate average placeholders per layout
Private Function GetAveragePlaceholders(ppt As Presentation) As Single
    On Error Resume Next
    Dim totalPlaceholders As Long
    Dim layoutCount As Long
    Dim layout As CustomLayout
    Dim shp As Shape
    
    totalPlaceholders = 0
    layoutCount = ppt.SlideMaster.CustomLayouts.Count
    
    If layoutCount = 0 Then
        GetAveragePlaceholders = 0
        Exit Function
    End If
    
    For Each layout In ppt.SlideMaster.CustomLayouts
        For Each shp In layout.Shapes
            If shp.Type = msoPlaceholder Then
                totalPlaceholders = totalPlaceholders + 1
            End If
        Next shp
    Next layout
    
    On Error GoTo 0
    GetAveragePlaceholders = CSng(totalPlaceholders) / CSng(layoutCount)
End Function

' --- Helper to join a collection of strings
Private Function JoinCollection(col As Collection, delimiter As String) As String
    Dim item As Variant
    Dim result As String
    Dim first As Boolean
    
    first = True
    For Each item In col
        If Not first Then
            result = result & delimiter
        End If
        result = result & item
        first = False
    Next item
    
    JoinCollection = result
End Function

' --- Enhanced placeholder type name function
Function GetPlaceholderTypeName(pType As Long) As String
    Select Case pType
        Case 1: GetPlaceholderTypeName = "Title"
        Case 2: GetPlaceholderTypeName = "Body"
        Case 3: GetPlaceholderTypeName = "CenterTitle"
        Case 4: GetPlaceholderTypeName = "Subtitle"
        Case 5: GetPlaceholderTypeName = "Date"
        Case 6: GetPlaceholderTypeName = "Footer"
        Case 7: GetPlaceholderTypeName = "Object"
        Case 8: GetPlaceholderTypeName = "Chart"
        Case 9: GetPlaceholderTypeName = "Table"
        Case 10: GetPlaceholderTypeName = "ClipArt"
        Case 11: GetPlaceholderTypeName = "OrgChart"
        Case 12: GetPlaceholderTypeName = "Media"
        Case 13: GetPlaceholderTypeName = "SlideImage"
        Case 14: GetPlaceholderTypeName = "Bitmap"
        Case 15: GetPlaceholderTypeName = "MediaClip"
        Case 16: GetPlaceholderTypeName = "SlideNumber"
        Case 17: GetPlaceholderTypeName = "Header"
        Case 18: GetPlaceholderTypeName = "Picture"
        Case 19: GetPlaceholderTypeName = "Content"
        Case Else: GetPlaceholderTypeName = "Unknown_Type_" & pType
    End Select
End Function