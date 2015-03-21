
Dim swApp               As SldWorks.SldWorks
Dim swModel             As SldWorks.ModelDoc2
Dim swExt               As SldWorks.ModelDocExtension
Dim featureMgr          As SldWorks.FeatureManager
Dim swSketchMgr         As SldWorks.SketchManager
Dim swSkRelMgr          As SldWorks.SketchRelationManager
Dim modelName           As String
Dim logFileName         As String
Dim indent              As String
Dim bullet              As String
Dim anchor_reg          As Collection



Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long


Public Enum swDocumentTypes_e
    swDocNONE = 0       '  Used to be TYPE_NONE
    swDocPART = 1       '  Used to be TYPE_PART
    swDocASSEMBLY = 2   '  Used to be TYPE_ASSEMBLY
    swDocDRAWING = 3    '  Used to be TYPE_DRAWING
End Enum

Sub main()
    Dim fileNumber          As Integer
    Dim imported            As Collection
    Dim import_chunk        As Variant
    Dim node_chunk          As Variant

    ' Initialize the anchor registry
    Set anchor_reg = New Collection

    ' Set globals
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    LastTLTag = ""
    modelName = swModel.GetTitle()
    logFileName = "C:\Users\nick\Dropbox\Documents\Projects - SMG\smgExport\" & modelName & ".smg"
    fileNumber = 2
    indent = "    "
    bullet = Left(indent, Len(indent) - 2) & "- "
    'bullet = "- "
    smg_TLKeys = Array("include", "asset global", "asset local")
    smgidname = "__smgID__"
    Dim header As String

    ' Open file for output
    Open logFileName For Input As #fileNumber

    Set imported = parse_file(fileNumber)

    ' We've got the imported file, now we need to manipulate it.
    ' (Temporarily) Reject the init chunk, and start construction with the
    ' main chunk.

    ' Look for first supported feature, then figure out base plane?
    ' 1) NEED TO PULL SKETCHES OUT AND PLACE BEFORE FEATURES!
    ' 2) For each feature in main,
    ' 3) 
    
    ' To be safe, we're going to create an entirely new document
    Dim templates       As Variant
    Dim template_name   As String
    templates = get_templates()
    template_name = templates(0)

    Set swModel = swApp.NewDocument(template_name, 0, 0, 0)
    Set swExt = swModel.Extension
    Set featureMgr = swModel.FeatureManager
    Set swSketchMgr = swModel.SketchManager

    ' Capture current user preferences and store them
    Dim dimvalsetting   As Boolean
    Dim graphicssetting As Boolean
    Dim dbsetting       As Boolean
    Dim modview         As SldWorks.ModelView
    set modview = swModel.ActiveView
    dimvalsetting = swApp.GetUserPreferenceToggle(swInputDimValOnCreate) 
    graphicssetting = modview.EnableGraphicsUpdate
    dbsetting = swSketchMgr.AddToDB

    ' Then disable graphics updating and dimension prompt
    swApp.SetUserPreferenceToggle swInputDimValOnCreate, False
    modview.EnableGraphicsUpdate = False
    swSketchMgr.AddToDB = True

    For Each import_chunk In imported
        If StrComp(import_chunk(0).Item("key"), "main") = 0 Then
            For Each node_chunk In import_chunk(1)
                build_node node_chunk
            Next node_chunk
        End If
    Next import_chunk

    ' Reset preferences
    swApp.SetUserPreferenceToggle swInputDimValOnCreate, dimvalsetting
    modview.EnableGraphicsUpdate = graphicssetting
    swSketchMgr.AddToDB = dbsetting

    ' Force a rebuild, then clear globals 
    Dim boolstatus  As Boolean
    boolstatus = swModel.ForceRebuild3(False)
    End
End Sub

Sub build_node(node As Variant)
    If StrComp(node(1).Item("type")(1), "sketch") = 0 Then
        build_sketch node
    End If
End Sub

Sub build_sketch(node As Variant)
    Dim boolstatus  As Boolean
    Dim parent      As String
    Dim entities    As Collection

    ' Don't forget that all the collections are wrapped.
    ' Pull some basic information about the sketch
    parent = node(1).Item("parent")(1)
    Set entities = node(1).Item("entities")(1)

    ' Select the base plane and open the sketch up
    boolstatus = swExt.SelectByID2(parent, "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    swSketchMgr.InsertSketch(True)

    ' Process the sketch entities
    insert_entities entities

    swSketchMgr.InsertSketch(False)
End Sub

Sub insert_entities(entities As Collection)
    Dim vContainer      As Variant
    Dim wrap_entity     As Variant
    Dim entity          As Collection

    For Each wrap_entity In entities
        ' Don't forget, everything is wrapped!
        Set entity = wrap_entity(1)

        ' Select cases for each type of entity
        Select Case entity.Item("type")(1)
            Case "point"
                create_point entity
            Case "line"
                create_line entity
            Case Else
                Debug.Print(entity.Item("def_sw")(1).Item("name")(1) & " import not yet supported.")
        End Select

    Next wrap_entity
End Sub

Function get_point_coords(location As String) As Variant
    Dim coords      As String
    Dim vCoordArr   As Variant
    Dim coordlen    As Integer
    Dim vref        As Variant
    Dim x           As Double
    Dim y           As Double
    Dim z           As Double
    Dim def_sw      As Collection
    
    ' Do whatever is needed to look for the correct point and return its coords
    ' Ensure a reference was passed (look for the asterisk)
    If StrComp(Left(location, 1), "*") = 0 Then
        ' Strip the reference of its beginning asterisk
        location = Right(location, Len(location) - 1)

        ' Resolve the reference
        vref = anchor_reg.Item(location)
        Set def_sw = vref(1).Item("def_sw")(1)
        coords = def_sw.Item("coordinates")(1)
    Else
        coords = location
    End If

    coordlen = Len(coords)

    ' Strip any leading/trailing brackets and whitespace
    If StrComp(Left(coords, 1), "[") = 0 Then
        coords = LTrim(Right(coords, coordlen - 1))
        coordlen = Len(coords)
    End If
    If StrComp(Right(coords, 1), "]") = 0 Then
        coords = RTrim(Left(coords, coordlen - 1))
        coordlen = Len(coords)
    End If
    ' Turn into an array
    vCoordArr = Split(coords, ", ")

    ' Convert to doubles
    x = CDbl(vCoordArr(0))
    y = CDbl(vCoordArr(1))
    z = CDbl(vCoordArr(2))

    get_point_coords = Array(x, y, z)
End Function

Sub create_point(entity As Collection)
    Dim id          As String
    Dim coords     As String
    Dim vCoordArr   As Variant
    Dim coordlen    As Integer
    Dim x           As Double
    Dim y           As Double
    Dim z           As Double
    Dim result      As Variant
    Dim def_sw      As Collection

    ' Get the solidworks definition collection
    Set def_sw = entity.Item("def_sw")(1)
    id = entity.Item("id")(1)

    ' Extract the coordinate string and then parse it to an array of doubles
    coords = def_sw.Item("coordinates")(1)
    vCoordArr = get_point_coords(coords)

    ' Extract x, y, z from vCoordArr
    x = vCoordArr(0)
    y = vCoordArr(1)
    z = vCoordArr(2)

    ' Conditionally create the point, if it's the correct solidworks type
    If def_sw.Item("type")(1) = 1 Then
        Set result = swSketchMgr.CreatePoint(x, y, z)
    End If

    anchor_reg.Item(id)(1).Item("def_sw")(1).Add result, "sw_entity"

End Sub

Sub create_line(entity As Collection)
    Dim id              As String
    Dim swSketchSegment As SldWorks.SketchSegment
    Dim swSketchLine    As SldWorks.SketchLine
    Dim startpt         As String
    Dim endpt           As String
    Dim startcoords     As Variant
    Dim endcoords       As Variant
    Dim result          As Variant
    Dim is_const        As Boolean
    Dim is_inf          As Boolean
    Dim def_sw          As Collection

    ' Get the solidworks def collection
    Set def_sw = entity.Item("def_sw")(1)
    id = entity.Item("id")(1)

    ' Extract the endpoints
    startpt = entity.Item("start")(1)
    endpt = entity.Item("end")(1)

    ' Get the parameters
    is_const = entity.Item("is_construction")(1)
    is_inf = entity.Item("is_infinite")(1)

    ' Get the associated coords
    startcoords = get_point_coords(startpt)
    endcoords = get_point_coords(endpt)

    ' Create the point
    Set result = swSketchMgr.CreateLine(startcoords(0), startcoords(1), startcoords(2), endcoords(0), endcoords(1), endcoords(2))
    Set swSketchSegment = result
    Set swSketchLine = result

    ' Set the parameters
    swSketchSegment.ConstructionGeometry = is_const

    If is_inf Then
        swSketchLine.MakeInfinite
    End If

    anchor_reg.Item(id)(1).Item("def_sw")(1).Add result, "sw_entity"

End Sub

Function get_templates() As Variant

    Dim part_temp   As String
    Dim assy_temp   As String
    
    part_temp = swApp.GetUserPreferenceStringValue(swDefaultTemplatePart)
    assy_temp = swApp.GetUserPreferenceStringValue(swDefaultTemplateAssembly)

    'part_temp = swApp.GetDocumentTemplate(swDocPART, "", 0, 0#, 0#)
    'assy_temp = swApp.GetDocumentTemplate(swDocASSEMBLY, "", 0, 0#, 0#)

    get_templates = Array(part_temp, assy_temp)

End Function

Function parse_file(fileNumber As Integer) As Collection
    Dim chunk_str       As String
    Dim chunk_coll      As Collection
    Dim line_coll       As Collection
    Dim has_whitespace  As Boolean
    Dim vchunk          As Variant
    Dim chunk           As String

    ' Initialize the chunk string
    chunk_str = ""
    Set chunk_coll = New Collection
    Set line_coll = New Collection

    ' Read the file until EOF and put into a collection of lines, ignoring comments and docstrings
    While Not EOF(fileNumber)
        ' Grab input line-by-line
        Line Input #fileNumber, line_str

        ' Ignore stream denotations
        If StrComp(Left(LTrim(line_str), 3), "---") = 0 Then
            Debug.Print(line_str)
        ' Ignore comments (trim left whitespace and compare to YAML comment #)
        ElseIf StrComp(Left(LTrim(line_str), 1), "#") = 0 Then
            Debug.Print(line_str)
        ' Not a comment, so let's worry about processing now
        Else
            line_coll.Add line_str
        End If
    Wend

    Set parse_file = parse_lines(line_coll)
End Function

' Take the collection of lines and turn it into a properly-nested collection
' If it has an anchor, the anchor should probably be used as the key.
Function parse_lines(lines As Collection) As Collection
    Dim parsed          As Collection
    Dim temp_coll       As Collection
    Dim control_coll    As Collection
    Dim line_str        As String
    Dim line            As Variant
    Dim indentlen       As Integer
    Dim splitt          As Variant
    Dim currentkey      As Variant
    Dim currentanc      As String
    Dim anchor          As String
    Dim vbuffer         As Variant
    Dim bullet          As Boolean
    Dim wrapper         As Variant
    Dim ii              As Integer
    Dim numlines        As Integer
    Dim has_anchor      As Boolean
    Dim parsed_temp     As Collection
    Dim valstr          As String
    Dim formed_buffer   As Variant
    Dim vtemp           As Variant

    'Set parse_lines = New Collection
    Set parsed = New Collection
    Set temp_coll = New Collection
    indentlen = 0
    numlines = lines.Count

    For ii = 1 To numlines
        line_str = lines.Item(ii)
        has_whitespace = (StrComp(LTrim(line_str), line_str) = 1)
        anchor = ""
        has_anchor = False
        bullet = False
        'wrapper = Array(New Collection, New Collection)


        ' If it's toplevel, either it has no whitespace, or we're at end of LIST.
        ' That means we need to output the previous key (if not first key),
        ' then continue (unless we're at the end of the LIST.)
        If Not has_whitespace Or ii = numlines Then
            ' Absolute first issue: is this the end of level, and this is not a toplevel key?
            If has_whitespace Then
                ' Add the line_str to the temporary collection
                ' DON'T FORGET TO REMOVE WHITESPACE BEFORE DOING THIS!
                ' Check to see if we've figured out how many characters of indentation to strip
                ' THIS CODE DUPLICATES BELOW
                If indentlen = 0 Then
                    indentlen = Len(line_str) - Len(LTrim(line_str))
                End If

                ' Remove the indentation and add the line to the temporary collection
                temp_coll.Add Right(line_str, Len(line_str) - indentlen)
            End If

            ' SANITY CHECK: 
            ' + in all cases, we currently have a fully-furnished temp_coll
            ' + if it exists, we should then parse the temp_coll
            ' + if we parse it, we need to add it to the buffer
            ' + in all cases, we need to form and output the buffer
            ' Now check that the temp collection is, in fact, empty.
            ' If it isn't, then we need to process it into sublevel collections.
            If temp_coll.Count > 0 Then
                ' First, process the temporary collection.
                Set parsed_temp = parse_lines(temp_coll)

                ' Now, add the parsed temp_coll to the buffer
                ' old_buffer = vbuffer
                vbuffer = Array(vbuffer(0), vbuffer(1), vbuffer(2), parsed_temp)
            End If

            ' Okay, now we just need to form and output the buffer.
            ' THIS IS CURRENTLY DUPLICATED FOR THE END OF LIST CASE'
            ' First, check if the value is a collection
            ' BUT FIRST make sure the buffer isn't empty.
            If Not IsEmpty(vbuffer) Then
                If TypeOf vbuffer(3) Is Collection Then
                    ' Clear the control_coll, if it exists
                    Set control_coll = New Collection

                    ' Check if anchored, and if so add both anchor and boolean
                    If StrComp(vbuffer(0), "") = 1 Then
                        control_coll.Add True, "has_anchor"
                        control_coll.Add vbuffer(0), "anchor"
                    Else
                        control_coll.Add False, "has_anchor"
                    End If

                    ' Add boolean for bulleting
                    control_coll.Add vbuffer(1), "has_bullet"

                    ' Add the item's key
                    control_coll.Add vbuffer(2), "key"

                    ' Finally, add this to the parsed
                    formed_buffer = Array(control_coll, vbuffer(3))
                    parsed.Add formed_buffer, formed_buffer(0).Item("key")
                    ' Check if it has an anchor, and if so, add it to the registry
                    If control_coll.Item("has_anchor") Then
                        anchor_reg.Add formed_buffer, formed_buffer(0).Item("key")
                    End IF
                ElseIf VarType(vbuffer(3)) = vbString Then
                    ' It's not a collection. Unfortunately, if it's anchored, byebye anchor.
                    ' Bullets are currently unhandled
                    ' Output to the parsed
                    formed_buffer = Array(vbuffer(2), vbuffer(3))
                    parsed.Add formed_buffer, formed_buffer(0)
                Else
                    Debug.Print("Unable to add buffer to parsed collection. ")
                End If
            End If

            ' Okay, we've done all of the outputting. Now we just need to clean up.
            ' Empty the temp_coll
            Set temp_coll = New Collection
            indentlen = 0


            ' First, create the toplevel key as a 2-member array and trim whitespace
            splitt = Split(line_str, ":", 2)
            splitt(0) = LTrim(splitt(0))
            splitt(1) = LTrim(splitt(1))

            ' Check for an anchor
            If StrComp(Left(splitt(1), 1), "&") = 0 Then
                anchor = LTrim(Right(splitt(1), Len(splitt(1)) - 1))
                ' If there's an anchor, don't treat it like a key.
                splitt(1) = ""
            Else
                anchor = ""
            End If

            ' Check for a bullet (can't be elseif because above is "value", this is "key")
            If StrComp(Left(splitt(0), 1), "-") = 0 Then
                bullet = True
                splitt(0) = LTrim(Right(splitt(0), Len(splitt(0)) - 1))
            Else
                bullet = False
            End If

            ' Okay, now we create the buffer. It goes: (anchor, bullet, key, value)
            ' value: splitt(1) NOTE THAT THIS MAY BE EMPTY
            ' key: splitt(0)
            vbuffer = Array(anchor, bullet, splitt(0), splitt(1))

            ' Okay, end of list, and we have no whitespace, so we need to output to parsed.
            ' Note that this cannot possibly be a collection.
            If Not has_whitespace And ii = numlines Then
                formed_buffer = Array(vbuffer(2), vbuffer(3))
                parsed.Add formed_buffer, formed_buffer(0)
            End If

        ' If there's no whitespace and not EOF, it's not a toplevel key.
        Else
            ' Add the line_str to the temporary collection
            ' DON'T FORGET TO REMOVE WHITESPACE BEFORE DOING THIS!
            ' Check to see if we've figured out how many characters of indentation to strip
            ' THIS CODE IS DUPLICATED FOR THE EOF CASE
            If indentlen = 0 Then
                indentlen = Len(line_str) - Len(LTrim(line_str))
            End If

            ' Remove the indentation and add the line to the temporary collection
            temp_coll.Add Right(line_str, Len(line_str) - indentlen)
        End If
    ' Move to the next line!
    Next ii

    'If has_anchor Then
    ''    anchor_reg.Add parsed, anchor
    'End If

    Set parse_lines = parsed
    
End Function

Function parse_nested_lines(lines As Collection) As Collection
    
End Function

' Turns the string of the toplevel bit into a properly nested collection
Function parse_chunk(chunk As String) As Collection
    ' Is it worthwhile to add a control collection?
    Dim chunkarr()  As String
    Dim lines       As Variant

    Set parse_chunk = New Collection

    chunkarr = Split(chunk, vbCrLf)
    lines = chunkarr

    For Each line In lines
        parse_chunk.Add line
    Next line

End Function
