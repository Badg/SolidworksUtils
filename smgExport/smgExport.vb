' ------------------------------------
'
' Notes:
'
' Ctrl+F "IMPENDING CHANGE"
' Store tree info in memory until ready to print?
'
' KNOWN BUGS:
' 
' Adding a feature wizard hole wizard thingajob results in failure to process 
' virtual/hidden/internal refplanes used for profile sketch 
'
'
' To output a subkey in the current implementation of build_chunk_dep, it needs to 
' have a key.  Format would be <indentlevel>, <key>, <value> still, so key...
' should exist at all times or there will be a stray :?  I suppose you could
' also output the value as the key, causing it to omit the : entirely.  That's
' a good strategy for something without a key, I guess.  In other cases, the 
' key should look like:
'		<\n><indent><key>: <value>
'		<\n><indent><key>: <value>
'		<\n><indent><key>: <value>
' etc
'




' Custom properties on a per-feature basis:
' http://help.solidworks.com/2012/English/api/sldworksapi/Add_and_Get_Custom_Properties_Example_VB.htm


' http://help.solidworks.com/2012/English/api/sldworksapi/Get_Plane_on_which_Sketch_Created_Example_VB.htm
' http://help.solidworks.com/2013/English/api/sldworksapi/Get_Faces_Associated_with_Feature_Example_VB.htm
' http://help.solidworks.com/2013/English/api/sldworksapi/Insert_Sketch_Text_and_Hole_Example_VB.htm
' swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swInputDimValOnCreate, False
' http://help.solidworks.com/2012/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IEntity_members.html
' http://help.solidworks.com/2012/English/api/sldworksapi/Get_All_Elements_of_Sketch_Example_VB.htm
' http://help.solidworks.com/2012/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IDimensionTolerance.html
' http://help.solidworks.com/2012/English/api/sldworksapi/Traverse_Feature_Dimensions_Example_VB.htm
' http://help.solidworks.com/2012/English/api/sldworksapi/Get_Display_Dimension_Properties_Example_VB.htm
' http://help.solidworks.com/2012/English/api/sldworksapi/Iterate_Through_Dimensions_in_Model_Example_VB.htm
' http://help.solidworks.com/2012/English/api/swdimxpertapi/Auto_Dimension_Scheme_Example_VB.htm
'
' -------------------------------------

' #####################################
' TODO
' ------
' I need a process entity sub in order to get the relation entities
' Should this have an explicit order, or be order-independent?
' Why the fuck can't I access dimensions?
' Tracking ID's are temporary in SW, but could be useful for internal script reference.
' If using them, assign tracking ID's on an as-needed basis, or always?  Since they're temporary, always might be easiest.
' Alternatively, persistent reference ID?
' Need to be able to equate a sketch edge with the body edge it creates
' Features and sketches are nonlinearly related.  For example, using a line in a rotational feature means that line corresponds
'    to a revolute surface; using it in an extruded feature makes it a planar face. Ultimately it should be the sketch entity
'    that any child features refer to, but how that sketch entity relates to the child entity is dependent upon the sketch
'    entity's relation to the feature.
' Need to track both faces and edges produced by features
' Named faces could be a route to automatically mate things together, have associativity between bolt patterns, etc
' https://forum.solidworks.com/message/332385

Option Explicit


' Declare globals
Dim swApp               As SldWorks.SldWorks
Dim swModel             As SldWorks.ModelDoc2
Dim featureMgr          As SldWorks.FeatureManager
Dim swSketchMgr         As SldWorks.SketchManager
Dim swSkRelMgr          As SldWorks.SketchRelationManager
Dim modelName           As String
Dim logFileName         As String
Dim fileNumber          As Integer
Dim ii                  As Integer
Dim indent				As String
Dim bullet				As String
Dim LastTLTag			As String
Dim out_coll            As New Collection
Dim smg_TLKeys          As Variant
Dim smgidname           As String
Dim smgIDs              As New Collection
Dim include_chunk       As Variant
Dim partdef             As New Collection
' These three are globals to allow for modification of any of them from
' any other of them
Dim vInclude            As Variant
Dim vGlobal             As Variant
Dim vLocal              As Variant

Public Enum swSketchRelationFilterType_e
    swAll = 0
    swDangling = 1
    swOverDefining = 2
    swExternal = 3
    swDefinedInContext = 4
    swLocked = 5
    swBroken = 6
    swSelectedEntities = 7
End Enum

Public Enum swConstraintType_e
    swConstraintType_INVALIDCTYPE = 0
    swConstraintType_DISTANCE = 1
    swConstraintType_ANGLE = 2
    swConstraintType_RADIUS = 3
    swConstraintType_HORIZONTAL = 4
    swConstraintType_VERTICAL = 5
    swConstraintType_TANGENT = 6
    swConstraintType_PARALLEL = 7
    swConstraintType_PERPENDICULAR = 8
    swConstraintType_COINCIDENT = 9
    swConstraintType_CONCENTRIC = 10
    swConstraintType_SYMMETRIC = 11
    swConstraintType_ATMIDDLE = 12
    swConstraintType_ATINTERSECT = 13
    swConstraintType_SAMELENGTH = 14
    swConstraintType_DIAMETER = 15
    swConstraintType_OFFSETEDGE = 16
    swConstraintType_FIXED = 17
    swConstraintType_ARCANG90 = 18
    swConstraintType_ARCANG180 = 19
    swConstraintType_ARCANG270 = 20
    swConstraintType_ARCANGTOP = 21
    swConstraintType_ARCANGBOTTOM = 22
    swConstraintType_ARCANGLEFT = 23
    swConstraintType_ARCANGRIGHT = 24
    swConstraintType_HORIZPOINTS = 25
    swConstraintType_VERTPOINTS = 26
    swConstraintType_COLINEAR = 27
    swConstraintType_CORADIAL = 28
    swConstraintType_SNAPGRID = 29
    swConstraintType_SNAPLENGTH = 30
    swConstraintType_SNAPANGLE = 31
    swConstraintType_USEEDGE = 32
    swConstraintType_ELLIPSEANG90 = 33
    swConstraintType_ELLIPSEANG180 = 34
    swConstraintType_ELLIPSEANG270 = 35
    swConstraintType_ELLIPSEANGTOP = 36
    swConstraintType_ELLIPSEANGBOTTOM = 37
    swConstraintType_ELLIPSEANGLEFT = 38
    swConstraintType_ELLIPSEANGRIGHT = 39
    swConstraintType_ATPIERCE = 40
    swConstraintType_DOUBLEDISTANCE = 41
    swConstraintType_MERGEPOINTS = 42
    swConstraintType_ANGLE3P = 43
    swConstraintType_ARCLENGTH = 44
    swConstraintType_NORMAL = 45
    swConstraintType_NORMALPOINTS = 46
    swConstraintType_SKETCHOFFSET = 47
    swConstraintType_ALONGX = 48
    swConstraintType_ALONGY = 49
    swConstraintType_ALONGZ = 50
    swConstraintType_ALONGXPOINTS = 51
    swConstraintType_ALONGYPOINTS = 52
    swConstraintType_ALONGZPOINTS = 53
    swConstraintType_PARALLELYZ = 54
    swConstraintType_PARALLELZX = 55
    swConstraintType_INTERSECTION = 56
    swConstraintType_PATTERNED = 57
    swConstraintType_ISOBYPOINT = 58
    swConstraintType_SAMEISOPARAM = 59
    swConstraintType_FITSPLINE = 60
End Enum

Public Enum swDimensionParamType_e 
	swDimensionParamTypeDoubleAngular = 1
	swDimensionParamTypeDoubleLinear = 0
	swDimensionParamTypeInteger = 2
	swDimensionParamTypeUnknown = -1
End Enum

Public Enum swSketchRelationEntityTypes_e
    swSketchRelationEntityType_Unknown = 0
    swSketchRelationEntityType_SubSketch = 1
    swSketchRelationEntityType_Point = 2
    swSketchRelationEntityType_Line = 3
    swSketchRelationEntityType_Arc = 4
    swSketchRelationEntityType_Ellipse = 5
    swSketchRelationEntityType_Parabola = 6
    swSketchRelationEntityType_Spline = 7
    swSketchRelationEntityType_Hatch = 8
    swSketchRelationEntityType_Text = 9
    swSketchRelationEntityType_Plane = 10
    swSketchRelationEntityType_Cylinder = 11
    swSketchRelationEntityType_Sphere = 12
    swSketchRelationEntityType_Surface = 13
    swSketchRelationEntityType_Dimension = 14
End Enum

Public Enum swInConfigurationOpts_e
    swConfigPropertySuppressFeatures = 0
    swThisConfiguration = 1
    swAllConfiguration = 2
    swSpecifyConfiguration = 3
End Enum

Public Enum swDimensionDrivenState_e
	swDimensionDriven = 1
	swDimensionDrivenUnknown = 0
	swDimensionDriving = 2
End Enum

Sub main()
    ' Set globals
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set featureMgr = swModel.FeatureManager
    Set swSketchMgr = swModel.SketchManager
	LastTLTag = ""
    modelName = swModel.GetTitle()
    logFileName = "C:\Users\nick\Dropbox\Documents\Projects - SMG\smgExport\" & modelName & ".smg"
    fileNumber = 1
	indent = "    "
	bullet = Left(indent, Len(indent) - 2) & "- "
    'bullet = "- "
    smg_TLKeys = Array("include", "asset global", "asset local")
    smgidname = "__smgID__"
    Dim header As String

    ' Open file for output
    Open logFileName For Output As fileNumber
    
    ' Give the file a header
    header = "--- !" & modelName & vbCrLf
    header = header & "# Source-Mutable Geometry (.smg) files are formatted as a subset of YAML."
    write_header header

    ' Build the output collection
    Dim tlkey As Variant
    Dim chunk As Variant
    Dim strKey As String
    For Each tlkey In smg_TLKeys
        strKey = tlkey
        chunk = Array(tlkey, New Collection)
        out_coll.Add chunk, tlkey
    Next tlkey

	' Run the (unimplemented) includes and add to partdef
    'vInclude = traverse_includes()
    'partdef.Add vInclude

    ' Create the init wrapped collection
    partdef.Add Array(New Collection, New Collection), "init"
    partdef.Item("init")(0).Add "init", "key"
	
    ' Set the root node
    Dim rootNode As SldWorks.TreeControlItem
    Dim treeNode As SldWorks.TreeControlItem
    Set rootNode = featureMgr.GetFeatureTreeRootItem()
    Dim nodeColl As Collection
    Dim node_vcoll   As Variant


    ' ######################################################################
    ' THIS WILL CREATE A WRAPPED COLLECTION FOR LOCAL. MIGHT MOVE TO NEW FUNC
    Dim local_coll        As Collection
    Dim local_control     As Collection

    Set local_coll = New Collection
    Set local_control = New Collection

    local_control.Add "main", "key"
    local_control.Add True, "bullet"
        
    ' Run the traversal, modifying out_coll
    If Not rootNode Is Nothing Then ' If the root node isn't empty
        Set treeNode = rootNode.GetFirstChild()
        While Not treeNode Is Nothing
            node_vcoll = traverse_tree(treeNode)
            If Not IsEmpty(node_vcoll) Then
                local_coll.Add node_vcoll, node_vcoll(0).Item("key")
                ' NEED TO ADD key for this based on container wrapper thing'
            End If
            Set treeNode = treeNode.GetNext()
        Wend
         ' Traverse the tree from the root node at the 0th level
    End If
    
    ' Define the local array and add it to partdef
    vLocal = Array(local_control, local_coll)
    ' THIS IS THE END OF THE WRAPPED LOCAL COLLECTION CODE
    ' ######################################################################

    ' ######################################################################
    ' Process the generated collection

    Dim currentkey As String
    For Each chunk In vLocal(1)
        ' First search for globals
        If Exists(chunk(0), "global") Then
            If chunk(0).Item("global") Then
                currentkey = chunk(0).Item("key")
                ' Add the chunk to the partdef at toplevel
                partdef.Item("init")(1).Add chunk, currentkey
                ' Remove the chunk from the local collection
                vLocal(1).Remove currentkey
            End If
        End If
    Next chunk

    return_sketch(vLocal)

    ' THIS IS THE END OF THE COLLECTION PROCESSING CODE
    ' ######################################################################

    ' Add local to the part definition
    partdef.Add vLocal
   
    ' Output the contents of the output_collection
    Dim chunkStr    As String
    For Each chunk In partdef
        chunkStr = compose_wrapped_collection(chunk)
        ' Print that shit
        Print #fileNumber, chunkStr
    Next chunk

    ' Finally, close the output file
    Close #fileNumber
    End
End Sub

' Extracts any sketches from parent features, returning them recursively.
' Need to make sure this isn't catching toplevel sketches that haven't been absorbed into another feature (wait, maybe not -- this doesn't appear to be an issue/)
' A HA! No, the above isn't an issue, because the key deletion/insertion is handled after the return statement. So, toplevel returns will just result in unused returns.
Function return_sketch(ByRef wr_coll As Variant) As Variant
    Dim thisitem            As Variant
    Dim thiskey             As String
    Dim thistype            As String
    Dim extracted_sketch    As Variant
    Dim extracted_reference As Variant
    Dim sketchkey           As String

    For Each thisitem In wr_coll(1)
        If TypeOf thisitem(1) Is Collection Then
            thiskey = thisitem(0).Item("key")

            If Exists(thisitem(1), "type") Then
                thistype = thisitem(1).Item("type")(1)

                If StrComp(thistype, "sketch") = 0 Then
                    return_sketch = thisitem
                Else
                    extracted_sketch = return_sketch(wr_coll(1).Item(thiskey))
                    If Not IsEmpty(extracted_sketch) Then
                        sketchkey = extracted_sketch(0).Item("key")
                        extracted_reference = Array("sketch", "*" & sketchkey)

                        wr_coll(1).Item(thiskey)(1).Add extracted_reference, extracted_reference(0), sketchkey
                        wr_coll(1).Item(thiskey)(1).Remove sketchkey
                        ' Add the extracted sketch to the wrapped collection, before the current item
                        wr_coll(1).Add extracted_sketch, sketchkey, thiskey
                    End If
                End If
            End If
        End If
    Next thisitem
End Function

Function Exists(ByVal coll As Collection, ByVal vKey As Variant) As Boolean
    On Error Resume Next
    coll.Item vKey
    Exists = (Err.Number = 0)
    Err.Clear
End Function

Sub write_header(str As String)
    Print #fileNumber, str
End Sub

Function block_format(block As String, Optional do_bullet As Boolean = False) As String
    If Len(block) > 0 Then
        block_format = Replace(block, vbCrLf, vbCrLf & indent)
        ' If the block doesn't start with a bullet point then give it an indent
        'If Not StrComp(Left(block_format, Len(bullet)), bullet) = 0 Then
        If do_bullet Then
            block_format = bullet & block_format
        Else
            block_format = indent & block_format
        End If
        'End If
        block_format = RTrim(block_format)
            block_format = vbCrLf & block_format
    End If
    ' block_format = block_format & vbCrLf
End Function

' Turns a collection (can be nested) into a YAML string representation thereof
Function compose_wrapped_collection(coll As Variant) As String
    Dim itm         As Variant
    Dim nested      As Collection
    Dim has_id      As Boolean
    Dim do_bullet   As Boolean
    Dim valstr      As String

    compose_wrapped_collection = ""

    If TypeOf coll Is Collection Then
        ' Does this ever even get used?
        For Each itm In coll
            compose_wrapped_collection = compose_wrapped_collection & itm(0) & ": " & itm(1) & vbCrLf
        Next itm
        compose_wrapped_collection = Left(compose_wrapped_collection, Len(compose_wrapped_collection) - 2)
    Else
        If TypeOf coll(1) Is Collection Then
            compose_wrapped_collection = compose_wrapped_collection & coll(0).Item("key") & ": "
            If Exists(coll(0), "anchor") Then
                If coll(0).Item("anchor") = True Then
                    compose_wrapped_collection = compose_wrapped_collection & "&" & coll(0).Item("key")
                End If
            End If
            If Exists(coll(0), "bullet") Then
                do_bullet = coll(0).Item("bullet")
            End If
            For Each itm In coll(1)
                compose_wrapped_collection = compose_wrapped_collection & block_format(compose_wrapped_collection(itm), do_bullet)
            Next itm
            
        ElseIf Len(coll(0)) > 0 Then
            ' It's a normal, bottom-level key, so add the key: value
            valstr = coll(1)
            ' First check for an anchor reference
            If StrComp(Left(valstr, 1), "*") = 0 Then
                ' Trim the reference
                valstr = Right(valstr, Len(valstr) - 1)
                ' Try to replace the name with the ID
                If Exists(smgIDs, valstr) Then
                    valstr = smgIDs.Item(valstr)
                End If

                ' Add the reference back
                valstr = "*" & valstr
            End If

            compose_wrapped_collection = compose_wrapped_collection & vbCrLf & coll(0) & ": " & valstr
        Else
            ' It's empty, so give the string some padding to keep the trim from erroring
            compose_wrapped_collection = "  " & compose_wrapped_collection
        End If
    End If
    If StrComp(Right(compose_wrapped_collection, 2), vbCrLf) = 0 Then
        compose_wrapped_collection = Left(compose_wrapped_collection, Len(compose_wrapped_collection) - 2)
    End If
    If StrComp(Left(compose_wrapped_collection, 2), vbCrLf) = 0 Then
        compose_wrapped_collection = Right(compose_wrapped_collection, Len(compose_wrapped_collection) - 2)
    End If
End Function

Function get_new_smgID(SW_name As String) As String
    Dim upperbound  As Long
    Dim lowerbound  As Integer
    Dim randInt     As Long
    Dim uniqueID    As Boolean
    Dim rowstr      As String
    Dim vrow        As Variant

    uniqueID = False
    upperbound = 90000000
    lowerbound = 0

    While Not uniqueID
        randInt = CLng((upperbound - lowerbound + 1) * Rnd()) + lowerbound
        get_new_smgID = Format(randInt, "00000000")
        uniqueID = True
        ' Well, this is stupid, but so is VBA
        For Each vrow In smgIDs
            rowstr = vrow
            If StrComp(rowstr, get_new_smgID) = 0 Then
                uniqueID = False
            End IF
        Next vrow
    Wend

    If Not Exists(smgIDs, SW_name) Then
        smgIDs.Add get_new_smgID, SW_name
    Else
        smgIDs.Add get_new_smgID
        ' get_new_smgID = smgIDs.Item(SW_name)
    End If
End Function

' Currently unimplemented, triggers file output of "include: __NONE__"
Function traverse_includes(Optional recursion As String = "__omit__") As Variant
	Dim includeStr	As String
    Dim includeArr  As Variant
    Dim control_coll    As Collection
    Dim def_coll        As Collection

    Set control_coll = New Collection
    Set def_coll = New Collection

    includeStr = ""
    includeArr = Array(includeStr, includeStr)

    control_coll.Add "include", "key"
    def_coll.Add includeArr, includeArr(0)

    traverse_includes = Array(control_coll, def_coll)
End Function

' traverse_tree is a recursive sub that will walk down the featuremanager tree and send each
' item there to be handled by its respective sub.
Function traverse_tree(node As SldWorks.TreeControlItem) As Variant
    ' Declare local variables
    Dim childNode       As SldWorks.TreeControlItem
    Dim siblingNode     As SldWorks.TreeControlItem
    Dim featureNode     As SldWorks.Feature
    '    Dim component      As SldWorks.Component2
    Dim childColl       As Collection
    Dim wrappedColl        As Variant
    Dim def_coll            As Collection
    Dim control_coll        As Collection
    Dim child               As Variant
    Dim container           As Variant
    ' Container is needed to hold traverse_tree, since otherwise you can't reference array members
	
    container = handle_treenode(node)

    ' Let's see if we will need to go a level deeper
    Set childNode = node.GetFirstChild()
    While Not childNode Is Nothing ' When any child node exists
        child = traverse_tree(childNode) ' Call recursively, at one more traverse level
        If Not IsEmpty(child) Then
            'If childColl.Count() > 0 Then
                If Not IsEmpty(container) Then
                    container(1).Add child, child(0).Item("key")
                Else
                    container = child
                End If
            'End If
        End If
        Set childNode = childNode.GetNext
    Wend

    traverse_tree = container
End Function

Function handle_treenode(node As SldWorks.TreeControlItem) As Variant
    Dim treeObjectType  As Long
    Dim treeObject      As Object
    Dim treeElement     As SldWorks.Feature
    Dim treeTag         As String
    Dim treeChunk       As Variant

    treeObjectType = node.ObjectType ' This node is something.  What is it?
    Set treeObject = node.Object
    
    If Not treeObject Is Nothing Then
        Select Case treeObjectType ' Figure out what the shit to do based on object type

            ' This is a feature.  Let's add it to the collection.
            Case SwConst.swTreeControlItemType_e.swFeatureManagerItem_Feature:
                ' Determine node type and handle respective things
                Set treeElement = treeObject
                Select Case (treeElement.GetTypeName2()) ' Decide what to do based on what the node actually is
                    ' It's a sketch:
                    Case "ProfileFeature" ' It's a sketch
                        handle_treenode = handle_sketch(treeElement.GetSpecificFeature2) ' Handle the sketch
                    ' It's a feature:
                    Case "BaseBody", "Blend", "BlendCut", "Boss", "BossThin", "Cavity", "Chamfer", "CirPattern", "CombineBodies", "CosmeticThread", "CurvePattern", "Cut", "CutThin", "Deform", "DeleteBody", "DelFace", "DerivedCirPattern", "DerivedLPattern", "Dome", "Draft", "Emboss", "Extrusion", "Fillet", "Helix", "HoleWzd", "Imported", "ICE", "LocalCirPattern", "LocalLPattern", "LPattern", "MirrorPattern", "MirrorSolid", "MoveCopyBody", "ReplaceFace", "RevCut", "Revolution", "RevolutionThin", "Shape", "Shell", "Split", "Stock", "Sweep", "SweepCut", "TablePattern", "Thicken", "ThickenCut", "VarFillet", "VolSweep", "VolSweepCut", "MirrorStock"
                        handle_treenode = handle_feature(treeElement)
                    
                    ' Reference geometry:
                    Case "CoordSys", "RefAxis", "ReferenceCurve", "RefPlane", "OriginProfileFeature", "RefPoint"
                        If Not IsEmpty(treeElement) Then
                            handle_treenode = define_entity(treeElement.GetSpecificFeature2)

                                ' http://help.solidworks.com/2013/English/api/swconst/SolidWorks.Interop.swconst~SolidWorks.Interop.swconst.swRefPlaneType_e.html
                                ' http://help.solidworks.com/2013/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IRefPlaneFeatureData_properties.html

                                ' http://help.solidworks.com/2013/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IRefAxis_members.html

                                ' http://help.solidworks.com/2013/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IRefPoint_members.html
                        Else 
                            Debug.Print("virtual feature is empty?")
                        End If
                    ' Okay, not a sketch, generic handling can resume:
                    Case Else ' It's anything other than a sketch
                        Debug.Print(treeElement.GetTypeName2() & " not yet supported for asset element export: " & treeElement.name)
                End Select

            ' This is something, and we don't know shit.
            Case Else:
                Debug.Print("Unhandled " & node.Text & " non-empty object.") ' Unhandled non-empty object.
        End Select
    Else ' "Is Nothing" object.  This currently handles the root node.
        Debug.Print("Unhandled " & node.Text & " empty node")
    End If
End Function

Function handle_feature(swFeat As SldWorks.Feature) As Variant
	Dim featureStr		As String
	Dim idStr			As String
    Dim vID             As Variant
    Dim strCoords       As String
    Dim kName           As Variant
    Dim sName           As String
    Dim kType           As Variant
    Dim kID             As Variant
    Dim kParent         As Variant
    Dim kCoords         As Variant
    Dim kDims           As Variant
    Dim kFaces          As Variant
    Dim def_coll    As Collection
    Dim control_coll    As Collection
    Dim kSMGID          As Variant
    Dim def_sw          As New Collection
    Dim con_sw          As New Collection
    Dim kdef_sw         As Variant
    
    Set def_coll = New Collection
    Set control_coll = New Collection

    con_sw.Add "def_sw", "key"

    sName = swFeat.name
    kName = Array("name", sName)
    def_sw.Add kName, kName(0)

    kSMGID = Array("id", get_new_smgID(sName))
    def_coll.Add kSMGID, kSMGID(0)
    control_coll.Add kSMGID(1), smgidname
    control_coll.Add kSMGID(1), "key"
	
    kType = Array("type", swFeat.GetTypeName2())
    def_coll.Add kType, kType(0)

    kDims = gather_dimensions(swFeat)
    ' Since dimensions sometimes returns empty (if a dimension belongs to something else, etc)
    If kDims(1).Count > 0 Then
        def_coll.Add kDims, kDims(0).Item("key")
    End If

    kFaces = gather_feat_faces(swFeat)
	def_coll.Add kFaces, kFaces(0).Item("key")

    kdef_sw = Array(con_sw, def_sw)
    def_coll.Add kdef_sw, kdef_sw(0).Item("key")

    handle_feature = Array(control_coll, def_coll)
End Function

' handle_sketch looks into the sketch, and outputs everything about it.
Function handle_sketch(swFeat As SldWorks.Feature) As Variant
	Dim swSketch			As SldWorks.Sketch
	Dim swParentEntity		As SldWorks.Entity
	Dim nEntType			As Long
    Dim vSketchSegments     As Variant
	Dim vSketchSegment		As Variant
	Dim vSketchPoints		As Variant
	Dim vSketchPoint		As Variant
	Dim vEntities			As Variant
    Dim restOfString        As String
	Dim name				As String
	Dim typ					As String
	Dim strParent			As String
	Dim sketchStr			As String
    Dim kName           As Variant
    Dim kType           As Variant
    Dim kID             As Variant
    Dim kParent         As Variant
    Dim kCoords         As Variant
    Dim kEnts               As Variant
    Dim kRels               As Variant
    Dim kDims               As Variant
    Dim def_coll            As Collection
    Dim control_coll        As Collection
    Dim kSMGID          As Variant
    Dim kAbsorbed       As Variant
    Dim vAbsorbed       As Variant
    Dim isAbsorbed      As String
    Dim sName           As String
    Dim def_sw          As New Collection
    Dim con_sw          As New Collection
    Dim kdef_sw         As Variant

    Set control_coll = New Collection
    Set def_coll = New Collection
	Set swSketch = swFeat

    con_sw.Add "def_sw", "key"

    sName = swFeat.name
    kName = Array("name", sName)
    def_sw.Add kName, kName(0)

	' Handle the usual common items and any attributes
    kSMGID = Array("id", get_new_smgID(sName))
    def_coll.Add kSMGID, kSMGID(0)
    control_coll.Add kSMGID(1), smgidname
    control_coll.Add kSMGID(1), "key"

    kType = Array("type", "sketch")
    def_coll.Add kType, kType(0)

    ' Figure out if the sketch is used in a feature
    vAbsorbed = swFeat.GetChildren
    If Not IsEmpty(vAbsorbed) Then
        isAbsorbed = "true"
    Else
        isAbsorbed = "false"
    End If
    kAbsorbed = Array("used_in_feature", isAbsorbed)
    def_sw.Add kAbsorbed, kAbsorbed(0)

    ' Handle sketch parent
	Set swParentEntity = swSketch.GetReferenceEntity(nEntType)
	kParent = Array("parent", define_entity(swParentEntity, False))
    def_coll.Add kParent, kParent(0)
	
	' Get sketch segments and points and join into single "entities" array
    vSketchSegments = swSketch.GetSketchSegments
	vSketchPoints = swSketch.GetSketchPoints2
	vEntities = Array(vSketchSegments, vSketchPoints)
		
    def_coll.Add gather_sketch_entities(vEntities)

    def_coll.Add gather_relations(swSketch)



	kDims = gather_dimensions(swFeat)
    def_coll.Add kDims, kDims(0).Item("key")

    control_coll.Add True, "anchor"

    kdef_sw = Array(con_sw, def_sw)
    def_coll.Add kdef_sw, kdef_sw(0).Item("key")

    handle_sketch = Array(control_coll, def_coll)
End Function

Function gather_sketch_entities(vEntities As Variant) As Variant
	Dim vEntContainer		As Variant
	Dim vEntity				As Variant
    Dim vDef                As Variant
    Dim kDef                As Variant
    Dim def_coll            As Collection
    Dim control_coll        As Collection

    Set control_coll = New Collection
    Set def_coll = New Collection

    control_coll.Add "entities", "key"
	
    ' For each entity, get and write its def_coll
    If (Not IsEmpty(vEntities)) Then
		For Each vEntContainer In vEntities
			For Each vEntity In vEntContainer
				If (Not IsEmpty(vEntity)) Then
                    ' Put the entity definition collection in def_coll
					kDef = define_entity(vEntity)
                    'Add the wrapped definition (<control coll, definition coll>) to the entity composure
					def_coll.Add kDef, kDef(0).Item("key")
				End If
			Next vEntity
		Next vEntContainer
	End If

    gather_sketch_entities = Array(control_coll, def_coll)
End Function

Function gather_dimensions(swFeat As SldWorks.Feature) As Variant
    Dim swDispDim   As SldWorks.DisplayDimension
    Dim kDef        As Variant
    Dim def_coll            As Collection
    Dim control_coll        As Collection

    Set control_coll = New Collection
    Set def_coll = New Collection

    control_coll.Add "dimensions", "key"

    ' Initialize output collection and get first dimension
	Set swDispDim = swFeat.GetFirstDisplayDimension

    ' Assuming it exists, 
	If Not swDispDim Is Nothing Then
		While (Not swDispDim Is Nothing)
			' If this dim belongs to this feature (and not an absorbed one)
			If swDispDim.GetDimension.GetFeatureOwner.name = swFeat.Name Then
				' Put the dimension definition collection in def_coll
				kDef = define_entity(swDispDim)
				def_coll.Add kDef, kDef(0).Item("key")
			End If
			Set swDispDim = swFeat.GetNextDisplayDimension(swDispDim)
		Wend
	End If

    gather_dimensions = Array(control_coll, def_coll)
End Function

Function gather_feat_faces(swFeat As SldWorks.Feature) As Variant
	Dim vFaces 		As Variant
	Dim vFace		As Variant
	Dim def_coll		As Collection
    Dim control_coll    As Collection
    Dim kDef        As Variant

    ' This is (temporarily?) needed to keep the faces from having the same collection key name
    Dim ii          As Integer

    ' Initialize collections
    Set control_coll = New Collection
    Set def_coll = New Collection

    ' Add key to control collection
    control_coll.Add "faces", "key"

    ' Get faces
	vFaces = swFeat.GetFaces
    ii = 0
	
	If Not IsEmpty(vFaces) Then
		For Each vFace in vFaces
            ' Put the faces collection in def_coll
			kDef = define_entity(vFace)
            def_coll.Add kDef, kDef(0).Item("key")
            ii = ii + 1
		Next vFace
	End If

    gather_feat_faces = Array(control_coll, def_coll)
End Function

Function gather_relations(swSketch As SldWorks.Sketch) As Variant
    Dim vSketchRels         As Variant
    Dim vSketchRel          As Variant
    Dim swSketchRel         As SldWorks.SketchRelation
    Dim relType             As Long
    Dim strEnt              As String
    Dim conText             As String
    Dim vEnts               As Variant
    Dim vEnt                As Variant
    Dim vEntTypes           As Variant
    Dim vEntType            As Variant
    Dim entCount            As Long
    Dim objEnt              As Object
    Dim vEntDef             As Variant
    Dim vDim 				As Variant
	Dim horizontals			As String
	Dim verticals			As String
	Dim isCollected			As Boolean
	Dim def_coll				As Collection
    Dim control_coll        As Collection
    Dim kDef                As Variant

    ' This is (temporarily?) needed to keep the faces from having the same collection key name
    Dim ii          As Integer

    ' Initialize output collection and get relation manager and relations
    Set def_coll = New Collection
    Set control_coll = New Collection
    Set swSkRelMgr = swSketch.RelationManager
    vSketchRels = swSkRelMgr.GetRelations(swAll)
    ii = 0

    ' Add key to control collection
    control_coll.Add "relations", "key"

    ' Figure out what kind of relation it is and then look at the entities therein
    If Not (IsEmpty(vSketchRels)) Then ' If there are relations...  Is this necessary?  Or absorbed in foreach?
        For Each vSketchRel In vSketchRels ' Then iterate over each relation
            ' Recast the variant as a sw.SketchRelation for access to methods and put the collection in def_coll
            Set swSketchRel = vSketchRel 
			kDef = define_entity(vSketchRel)
            def_coll.Add kDef, kDef(0).Item("key")
            ii = ii + 1
		Next vSketchRel

        control_coll.Add Array("relations", "key")
        gather_relations = Array(control_coll, def_coll)
    End If
End Function

' define_entity takes a solidworks sketchSegment, sketchPoint, relation entity, or entity
' and returns a name in the form "[type] [id] @ [location]" and, by default, its defining
' characteristics (ex line endpoints, point coords, etc)
Function define_entity(vEntity As Variant, Optional getDef As Boolean = True) As Variant
    Dim def_coll           As Collection
    Dim control_coll         As Collection
    Dim swSketchText        As SldWorks.SketchText
    Dim swSketchPoint       As SldWorks.SketchPoint
    Dim swSketchSegment     As SldWorks.SketchSegment
    Dim swSketchLine        As SldWorks.SketchLine ' Create a line object
    Dim swSketchEllipse     As SldWorks.SketchEllipse
    Dim swSketchArc         As SldWorks.SketchArc
    Dim swSketchParabola    As SldWorks.SketchParabola
    Dim swSketchSpline      As SldWorks.SketchSpline
	Dim swRefPlane			As SldWorks.refplane
	Dim swFace				As SldWorks.face2
	Dim swDispDim			As SldWorks.DisplayDimension
	Dim swSketchRelation	As SldWorks.SketchRelation
	Dim swEdge				As SldWorks.Edge
	Dim swFeat				As SldWorks.Feature
	Dim swRefPt				As SldWorks.RefPoint
    Dim kSMGID          As Variant
    Dim def_name            As String
	
    Set control_coll = New Collection

    ' Recast sketchsegments as their actual entities and re-call function recursively.  This may not actually get used; oh well.
    ' Specific entities must come first, or the broader type will execute before the narrow one, resulting in an infinite recursion.
    If TypeOf vEntity Is SldWorks.SketchPoint Then
		Set swSketchPoint = vEntity
        Set def_coll = explode_point(swSketchPoint)
        
    ElseIf TypeOf vEntity Is SldWorks.SketchLine Then
		Set swSketchLine = vEntity
        Set def_coll = explode_line(swSketchLine, getDef)
        
    ElseIf TypeOf vEntity Is SldWorks.SketchArc Then
		Set swSketchArc = vEntity
        Set def_coll = explode_arc(swSketchArc)
		
    ElseIf TypeOf vEntity Is SldWorks.RefPlane Then
		' Process refplane
		Set swRefPlane = vEntity
		Set def_coll = explode_refplane(swRefPlane)
        ' NODETYPE declaration
        If def_coll.Item("def_sw")(1).Item("type2")(1) = 10 Then
            control_coll.Add "global", "TLtag"
            control_coll.Add True, "global"
        'Else
            'TLTag = "asset local"
        End If
		
	ElseIf TypeOf vEntity Is SldWorks.face2 Then
		' Process face
		Set swFace = vEntity
		Set def_coll = explode_face(swFace)
	
	ElseIf TypeOf vEntity Is SldWorks.DisplayDimension Then
		Set swDispDim = vEntity
		Set def_coll = explode_dimension(swDispDim)
		
	ElseIf TypeOf vEntity Is SldWorks.SketchRelation Then
		Set swSketchRelation = vEntity
		Set def_coll = explode_relation(swSketchRelation)
		
	ElseIf TypeOf vEntity Is SldWorks.Edge Then
		Set swEdge = vEntity
		Set def_coll = explode_edge(swEdge)
		
	ElseIf TypeOf vEntity Is SldWorks.RefPoint Then
		Set swRefPt = vEntity
		Set def_coll = explode_refpt(swRefPt)
            
	Else ' type of entity not yet supported, etc
		Set swFeat = vEntity
        Dim swType As String
        swType = swFeat.GetTypeName

        If StrComp(swType, "OriginProfileFeature") = 0 Then
            Set def_coll = explode_origin(swFeat)
            control_coll.Add "global", "TLTag"
            control_coll.Add True, "global"
        Else
            Debug.Print("Type of entity not yet definable: " & swFeat.Name & " of type " & TypeName(swFeat)) & " / " & swFeat.GetTypeName
            Set def_coll = New Collection
            def_coll.Add Array("name", swFeat.Name), "name"
            def_coll.Add Array("type", swType), "type"
            def_coll.Add Array("unsupported", "True"), "unsupported"
        End If

    End If
        
    If getDef Then ' If ( by default or otherwise) we're getting the definition as well
        If Not Exists(control_coll, smgidname) Then   
            If Exists(def_coll, "name")  Then
                def_name = def_coll.Item("name")(1)
            Else
                def_name = def_coll.Item("def_sw")(1).Item("name")(1)
            End If
            kSMGID = Array("id", get_new_smgID(def_name))
            def_coll.Add kSMGID, kSMGID(0)
            control_coll.Add kSMGID(1), smgidname
        End If
        If Not Exists(def_coll, "key") Then
            control_coll.Add control_coll.Item(smgidname), "key"
        End If
        control_coll.Add True, "anchor"
        define_entity = Array(control_coll, def_coll)
    Else
		' Output only the entity's well-formed name as a reference
		' IMPENDING CHANGE fix this to use collection keys
        If Exists(control_coll, smgidname) Then
            define_entity = "*" & control_coll.Item(smgidname)(1)
        Else
            If Exists(def_coll, "name") Then
                define_entity = "*" & def_coll.Item("name")(1)
            Else
                define_entity = "*" & def_coll.Item("def_sw")(1).Item("name")(1)
            End If
        End If
    End If
End Function

Function explode_origin(swFeat As SldWorks.Feature) As Collection
    Dim TLTag           As String
    Dim kName           As Variant
    Dim kType           As Variant
    Dim kType2          As Variant
    Dim kTLTag          As Variant

    Dim def_sw          As New Collection
    Dim con_sw          As New Collection
    Dim kdef_sw         As Variant

    con_sw.Add "def_sw", "key"
    
    Set explode_origin = New Collection
    
    ' Name field
    kName = Array("name", swFeat.Name)
    def_sw.Add kName, kName(0)
    
    ' Type field
    kType = Array("type", "origin")
    explode_origin.Add kType, kType(0)

    ' def_sw field
    kdef_sw = Array(con_sw, def_sw)
    explode_origin.Add kdef_sw, kdef_sw(0).Item("key")
End Function

Function explode_dimension(swDispDim As SldWorks.DisplayDimension) As Collection
	Dim swDim 		As SldWorks.Dimension
	Dim swAnno		As SldWorks.Annotation
	Dim vEnts		As Variant
	Dim vEnt		As Variant
    Dim dimName     As String
    Dim dimValue    As Double
	Dim dimType		As String
	Dim dimState	As String
	Dim dimParent	As String
    Dim strMembers	As String
	
	Dim strAtt		As String
	Dim kName		As Variant
	Dim kType 		As Variant
	Dim kAtt		As Variant
	Dim kState		As Variant
	Dim kParent		As Variant
	Dim kValue		As Variant
    Dim def_sw          As New Collection
    Dim con_sw          As New Collection
    Dim kdef_sw         As Variant

    con_sw.Add "def_sw", "key"
	
	Set explode_dimension = New Collection
	
	' Get the associated dimension object and annotation object
    Set swDim = swDispDim.GetDimension
    Set swAnno = swDispDim.GetAnnotation
    vEnts = swAnno.GetAttachedEntities3
	
	' Get the name
	kName = Array("name", swDim.Name)
	def_sw.Add kName, kName(0)
	
	' Get the parent
	kParent = Array("parent", swDim.GetFeatureOwner.name)
	explode_dimension.Add kParent, kParent(0)
	
	' Get the user value
	kValue = Array("value", CStr(swDim.GetUserValueIn(swModel)))
	explode_dimension.Add kValue, kValue(0)
	
	' Get dimension type
	Select Case (swDispDim.Type2)
	   Case swDimensionType_e.swOrdinateDimension
		   kType = Array("type", "ordinate")
	   Case swDimensionType_e.swLinearDimension
		   kType = Array("type", "linear")
	   Case swDimensionType_e.swAngularDimension
		   kType = Array("type", "angular")
	   Case swDimensionType_e.swArcLengthDimension
		   kType = Array("type", "arclength")
	   Case swDimensionType_e.swRadialDimension
		   kType = Array("type", "radial")
	   Case swDimensionType_e.swDiameterDimension
		   kType = Array("type", "diameter")
	   Case swDimensionType_e.swHorOrdinateDimension
		   kType = Array("type", "horizontalordinate")
	   Case swDimensionType_e.swVertOrdinateDimension
		   kType = Array("type", "verticalordinate")
	   Case swDimensionType_e.swZAxisDimension
		   kType = Array("type", "zaxis")
	   Case swDimensionType_e.swChamferDimension
		   kType = Array("type", "chamfer")
	   Case swDimensionType_e.swHorLinearDimension
		   kType = Array("type", "horizontallinear")
	   Case swDimensionType_e.swVertLinearDimension
		   kType = Array("type", "verticallinear")
	   Case swDimensionType_e.swScalarDimension
		   kType = Array("type", "scalar")
	   Case Else
		   kType = Array("type", "unknown")
	End Select
	explode_dimension.Add kType, kType(0)
	
	' Handle any attached entities
    If (Not IsEmpty(vEnts)) Then
		strAtt = "["
        For Each vEnt In vEnts
			strAtt = strAtt & define_entity(vEnt, False) & ", "
        Next vEnt
		' Remove extra ", " and close bracket
		strAtt = Left(strAtt, Len(strAtt) - 2) & "]"
		' Bind to key
		kAtt = Array("attachments", strAtt)
		explode_dimension.Add kAtt, kAtt(0)
    End If
	
	' Get dimension state
	Select Case (swDim.DrivenState)
		Case swDimensionDriven
			kState = Array("state", "driven")
		Case swDimensionDrivenUnknown
			kState = Array("state", "unknown")
		Case swDimensionDriving
			kState = Array("state", "driving")
	End Select
	explode_dimension.Add kState, kState(0)

    kdef_sw = Array(con_sw, def_sw)
    explode_dimension.Add kdef_sw, kdef_sw(0).Item("key")
End Function

Function explode_point(swSketchPoint As SldWorks.SketchPoint) As Collection
	' Automatically re-casts object during argument passing. 
	' "kEyed" variant arrays are prefaced as such.
    Dim swSketch    	As SldWorks.Sketch
    Dim swFeat     		As SldWorks.Feature
	Dim vID				As Variant
	Dim strCoords		As String
	Dim kName			As Variant
	Dim kType			As Variant
	Dim kID				As Variant
	Dim kParent			As Variant
	Dim kCoords			As Variant
    Dim kSWType         As Variant
    Dim def_sw          As New Collection
    Dim con_sw          As New Collection
    Dim kdef_sw         As Variant
	
	Set explode_point = New Collection

    con_sw.Add "def_sw", "key"
	
	' Type field
	kType = Array("type", "point")
	explode_point.Add kType, kType(0)
	
	' ID field
	vID = swSketchPoint.GetID
	kID = Array("id", vID(0) & "-" & vID(1))
	def_sw.Add kID, kID(0)
    
	' Parent field
	Set swSketch = swSketchPoint.GetSketch
	Set swFeat = swSketch
	kParent = Array("parent", swFeat.Name)
	explode_point.Add kParent, kParent(0)

    kSWType = Array("type", swSketchPoint.Type)
    def_sw.Add kSWType, kSWType(0)
	
	' Name field
	kName = Array("name", "point" & vID(1) & "@" & kParent(1))
	def_sw.Add kName, kName(0)
	
	' Coords field
	strCoords = "[" & Str(swSketchPoint.X) & ", " & Str(swSketchPoint.Y) & ", " & Str(swSketchPoint.Z) & "]"
	kCoords = Array("coordinates", strCoords)
	def_sw.Add kCoords, kCoords(0)

    kdef_sw = Array(con_sw, def_sw)
    explode_point.Add kdef_sw, kdef_sw(0).Item("key")
End Function

Function explode_line(swSketchLine As SldWorks.SketchLine, Optional do_fulldef As Boolean = True) As Collection
	' Extracts definition information from a sketch line.
	' "kEyed" variant arrays are prefaced as such.
	' I should probably have this check if it's infinite at some point.
	' Should these points reference their respective points?
	Dim swSketchSegment	As SldWorks.SketchSegment
    Dim swFeat     		As SldWorks.Feature
	Dim vID 			As Variant
	Dim kName			As Variant
	Dim kType			As Variant
	Dim kID				As Variant
	Dim kParent			As Variant
	Dim kStart			As Variant
	Dim kEnd			As Variant
    Dim kConstruct      As Variant
    Dim kInf            As Variant
    Dim def_sw          As New Collection
    Dim con_sw          As New Collection
    Dim kdef_sw         As Variant
    Dim kConstraints    As Variant
    Dim constraint_coll As New Collection
    Dim constraint_con  As New Collection
    Dim rel             As Variant
    Dim def_rel         As Variant
    Dim vrels           As Variant
	
	Set explode_line = New Collection
    Set swSketchSegment = swSketchLine
    Set swFeat = swSketchSegment.GetSketch

    con_sw.Add "def_sw", "key"
	
	' Type field
	kType = Array("type", "line")
	explode_line.Add kType, kType(0)
    
    ' Parent field
    kParent = Array("parent", swFeat.Name)
    explode_line.Add kParent, kParent(0)
    
    ' ID field
    vID = swSketchLine.GetID
    kID = Array("id", vID(0) & "-" & vID(1))
    def_sw.Add kID, kID(0)
    
    ' Name field
    kName = Array("name", "line" & vID(1) & "@" & kParent(1))
    def_sw.Add kName, kName(0)

    ' Check if infinite
    kInf = Array("is_infinite", swSketchLine.Infinite)
    explode_line.Add kInf, kInf(0)
	
    ' Check if construction geometry
    kConstruct = Array("is_construction", swSketchSegment.ConstructionGeometry)
    explode_line.Add kConstruct, kConstruct(0)

	' Start point field
	kStart = Array("start", define_entity(swSketchLine.GetStartPoint2, False))
	explode_line.Add kStart, kStart(0)
	
	' End point field
	kEnd = Array("end", define_entity(swSketchLine.GetEndPoint2, False))
	explode_line.Add kEnd, kEnd(0)

    ' This is needed to prevent infinite recursion
    If do_fulldef Then
        ' Pull the relations
        constraint_con.Add "constraints", "key"
        vrels = swSketchSegment.GetRelations
        For Each rel in vrels
            def_rel = define_entity(rel)
            constraint_coll.Add def_rel, def_rel(0).Item("key")
        Next rel
        kConstraints = Array(constraint_con, constraint_coll)
        explode_line.Add kConstraints, kConstraints(0).Item("key")
    End If

    ' Solidworks definition field
    kdef_sw = Array(con_sw, def_sw)
    explode_line.Add kdef_sw, kdef_sw(0).Item("key")
End Function

Function explode_arc(swSketchArc As SldWorks.SketchArc) As Collection
	' Extracts definition information from a sketch arc.
	Dim swSketchSegment	As SldWorks.SketchSegment
    Dim swFeat     		As SldWorks.Feature
	Dim vID 			As Variant
	Dim kName			As Variant
	Dim kType			As Variant
	Dim kID				As Variant
	Dim kParent			As Variant
	Dim kStart			As Variant
	Dim kEnd			As Variant
	Dim kCenter			As Variant
    Dim kConstruct      As Variant
    Dim def_sw          As New Collection
    Dim con_sw          As New Collection
    Dim kdef_sw         As Variant
	
	Set explode_arc = New Collection 
    Set swSketchSegment = swSketchArc
    Set swFeat = swSketchSegment.GetSketch

    con_sw.Add "def_sw", "key"

	' ID field
	vID = swSketchArc.GetID
	kID = Array("id", vID(0) & "-" & vID(1))
	def_sw.Add kID, kID(0)
	
	' Center point field
	kCenter = Array("center", define_entity(swSketchArc.GetCenterPoint2, False))
	explode_arc.Add kCenter, kCenter(0)
	
	' Parent field
	kParent = Array("parent", swFeat.Name)
	explode_arc.Add kParent, kParent(0)
	
	' Name field
	kName = Array("name", "arc" & vID(1) & "@" & kParent(1))
	def_sw.Add kName, kName(0)
    
    ' Check if construction geometry
    kConstruct = Array("is_construction", swSketchSegment.ConstructionGeometry)
    explode_arc.Add kConstruct, kConstruct(0)
	
	If (Not swSketchArc.IsCircle = 1) Then   ' If it's an arc and not a circle
		' Type field
		kType = Array("type", "arc")
			
		' Start point field
		kStart = Array("start", define_entity(swSketchArc.GetStartPoint2, False))
		explode_arc.Add kStart, kStart(0)
		
		' End point field
		kEnd = Array("end", define_entity(swSketchArc.GetEndPoint2, False))
		explode_arc.Add kEnd, kEnd(0)
	Else ' Then it's a circle
		' Type field
		kType = Array("type", "circle")
	End If	
	' Finally, add the type field to the collection
	explode_arc.Add kType, kType(0)

    ' Add the sw_def field
    kdef_sw = Array(con_sw, def_sw)
    explode_arc.Add kdef_sw, kdef_sw(0).Item("key")
End Function
    
Function explode_refplane(swRefPlane As SldWorks.refplane) As Collection
    Dim swSketch    	As SldWorks.Sketch
    Dim swFeat     		As SldWorks.Feature
	Dim swInfo			As SldWorks.RefPlaneFeatureData
	Dim TLTag			As String
	Dim kName			As Variant
	Dim kType			As Variant
	Dim kType2			As Variant
	Dim kTLTag			As Variant
    Dim def_sw          As New Collection
    Dim con_sw          As New Collection
    Dim kdef_sw         As Variant
	
	Set explode_refplane = New Collection
	Set swFeat = swRefPlane
	Set swInfo = swFeat.GetDefinition

    con_sw.Add "def_sw", "key"
	
	' Name field
	kName = Array("name", swFeat.Name)
	def_sw.Add kName, kName(0)
    
    ' Type2 field
    kType2 = Array("type2", swInfo.Type2)
    def_sw.Add kType2, kType2(0)
	
	' Type field
    If kType2(1) = 10 Then
        kType = Array("type", "baseplane")
    Else
        kType = Array("type", "refplane")
    End If
    explode_refplane.Add kType, kType(0)

    ' Add the sw_def field
    kdef_sw = Array(con_sw, def_sw)
    explode_refplane.Add kdef_sw, kdef_sw(0).Item("key")
End Function

Function explode_refpt(swRefPt As SldWorks.RefPoint) As Collection
	Dim swInfo			As SldWorks.RefPointFeatureData
    Dim swSketch    	As SldWorks.Sketch
    Dim swFeat     		As SldWorks.Feature
	Dim TLTag			As String
	Dim kName			As Variant
	Dim kType			As Variant
	Dim kType2			As Variant
	Dim kTLTag			As Variant
    Dim def_sw          As New Collection
    Dim con_sw          As New Collection
    Dim kdef_sw         As Variant
	
	Set explode_refpt = New Collection
	Set swFeat = swRefPt
	Set swInfo = swFeat.GetDefinition

    con_sw.Add "def_sw", "key"
	
	' Name field
	kName = Array("name", swFeat.Name)
	def_sw.Add kName, kName(0)
	
	' Type field
	kType = Array("type", "refpoint")
	explode_refpt.Add kType, kType(0)
	
	' Type2 field
	kType2 = Array("type2", swInfo.Type)
	def_sw.Add kType2, kType2(0)
	
	' NODETYPE declaration
	'If kType2(1) = 10 Then
	''	TLTag = "asset global"
	'Else
	''	TLTag = "asset local"
	'End If
	'kTLTag = Array("NODETYPE", TLTag)
	'explode_refpt.Add kTLTag, kTLTag(0)

    ' Add the sw_def field
    kdef_sw = Array(con_sw, def_sw)
    explode_refpt.Add kdef_sw, kdef_sw(0).Item("key")
End Function

Function explode_face(swFace As SldWorks.face2) As Collection
    Dim swEntity        As SldWorks.Entity
	Dim vNorm			As Variant
	Dim strNorm			As String
	Dim kName			As Variant
	Dim kType			As Variant
	Dim kNorm			As Variant
	Dim kFeature		As Variant
    Dim def_sw          As New Collection
    Dim con_sw          As New Collection
    Dim kdef_sw         As Variant
	
	Set explode_face = New Collection
    Set swEntity = swFace

    con_sw.Add "def_sw", "key"
	
	vNorm = swFace.Normal
	strNorm = "[" & Str(vNorm(0)) & ", " & Str(vNorm(1)) + ", " + Str(vNorm(2)) + "]"
	
	' Name field
	kName = Array("name", swEntity.ModelName)
	def_sw.Add kName, kName(0)
	
	' Type field
	kType = Array("type", "face")
	explode_face.Add kType, kType(0)
	
	' Norm field
	kNorm = Array("normal vector", strNorm)
	def_sw.Add kNorm, kNorm(0)

    ' Add the sw_def field
    kdef_sw = Array(con_sw, def_sw)
    explode_face.Add kdef_sw, kdef_sw(0).Item("key")
End Function

Function explode_edge(swEdge As SldWorks.Edge) As Collection
	Dim kName			As Variant
    Dim def_sw          As New Collection
    Dim con_sw          As New Collection
    Dim kdef_sw         As Variant
	
	Set explode_edge = New Collection

    con_sw.Add "def_sw", "key"
	
	' Name field
	kName = Array("name", "anonymous edge")
	def_sw.Add kName, kName(0)

    ' Add the sw_def field
    kdef_sw = Array(con_sw, def_sw)
    explode_edge.Add kdef_sw, kdef_sw(0).Item("key")
End Function

Function explode_relation(swSketchRel As SldWorks.SketchRelation) As Collection
    Dim swSketch    	As SldWorks.Sketch
    Dim swFeat     		As SldWorks.Feature
    Dim swEntity        As SldWorks.Entity
	Dim vEnts			As Variant
	Dim vEnt			As Variant
	Dim strType			As String
	Dim intEnts			As Integer
	Dim kName			As Variant
	Dim kType			As Variant
	Dim kEnt1			As Variant
	Dim kEnt2			As Variant
	Dim kEnt3			As Variant
    Dim def_sw          As New Collection
    Dim con_sw          As New Collection
    Dim kdef_sw         As Variant
	
	Set explode_relation = New Collection
    ' Set swFeat = swSketchRel

    con_sw.Add "def_sw", "key"
	
	' Name field
    'kName = Array("name", swFeat.Name)
	kName = Array("name", "anonymous relation")
	def_sw.Add kName, kName(0)
	
	' Get the relation's associated entities
	vEnts = swSketchRel.GetDefinitionEntities
	
	' Note: source: global is a placeholder for referencing the correct dimension
	Select Case (swSketchRel.GetRelationType())
        '          Case swConstraintType_INVALIDCTYPE
        '               strType = "invalid"
		Case swConstraintType_DISTANCE
			strType = "distance"
			kEnt1 = Array("source", "global")
			kEnt2 = Array("destination1", define_entity(vEnts(0), False))
			kEnt3 = Array("destination2", define_entity(vEnts(1), False))
		Case swConstraintType_ANGLE
			strType = "angle"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_RADIUS
			strType = "radius"
			kEnt1 = Array("source", "global")
			kEnt2 = Array("destination", define_entity(vEnts(0), False))
		Case swConstraintType_HORIZONTAL
			strType = "horizontal"
			kEnt1 = Array("source", "global")
			kEnt2 = Array("destination", define_entity(vEnts(0), False))
		Case swConstraintType_VERTICAL
			strType = "vertical"
			kEnt1 = Array("source", "global")
			kEnt2 = Array("destination", define_entity(vEnts(0), False))
		Case swConstraintType_TANGENT
			strType = "tangent"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_PARALLEL
			strType = "parallel"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_PERPENDICULAR
			strType = "perpendicular"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_COINCIDENT
			strType = "coincident"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_CONCENTRIC
			strType = "concentric"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_SYMMETRIC
			strType = "symmetric"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination1", define_entity(vEnts(1), False))
			kEnt3 = Array("destination2", define_entity(vEnts(2), False))
		Case swConstraintType_ATMIDDLE
			strType = "at midpoint"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_ATINTERSECT
			strType = "at intersection"
			kEnt1 = Array("source1", define_entity(vEnts(0), False))
			kEnt2 = Array("source2", define_entity(vEnts(1), False))
			kEnt3 = Array("destination", define_entity(vEnts(2), False))
		Case swConstraintType_SAMELENGTH
			strType = "equal length"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_DIAMETER
			strType = "diameter"
			kEnt1 = Array("source", "global")
			kEnt2 = Array("destination", define_entity(vEnts(0), False))
		Case swConstraintType_OFFSETEDGE
			strType = "offset from edge"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_FIXED
			strType = "fixed"
			kEnt1 = Array("source", "global")
			kEnt2 = Array("destination", define_entity(vEnts(0), False))
        '            Case swConstraintType_ARCANG90
        '            Case swConstraintType_ARCANG180
        '            Case swConstraintType_ARCANG270
        '            Case swConstraintType_ARCANGTOP
        '            Case swConstraintType_ARCANGBOTTOM
        '            Case swConstraintType_ARCANGLEFT
        '            Case swConstraintType_ARCANGRIGHT
		Case swConstraintType_HORIZPOINTS
			strType = "horizontal from point"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_VERTPOINTS
			strType = "vertical from point"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_COLINEAR
			strType = "colinear"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_CORADIAL
			strType = "coradial"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
        '            Case swConstraintType_SNAPGRID
        '            Case swConstraintType_SNAPLENGTH
        '            Case swConstraintType_SNAPANGLE
        '            Case swConstraintType_USEEDGE
        '            Case swConstraintType_ELLIPSEANG90
        '            Case swConstraintType_ELLIPSEANG180
        '            Case swConstraintType_ELLIPSEANG270
        '            Case swConstraintType_ELLIPSEANGTOP
        '            Case swConstraintType_ELLIPSEANGBOTTOM
        '            Case swConstraintType_ELLIPSEANGLEFT
        '            Case swConstraintType_ELLIPSEANGRIGHT
		Case swConstraintType_ATPIERCE
			strType = "pierce"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_DOUBLEDISTANCE
			strType = "doubled distance"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
		Case swConstraintType_MERGEPOINTS
			strType = "merged points"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination", define_entity(vEnts(1), False))
        '            Case swConstraintType_ANGLE3P
        '            Case swConstraintType_ARCLENGTH
        '            Case swConstraintType_NORMAL
        '            Case swConstraintType_NORMALPOINTS
        '            Case swConstraintType_SKETCHOFFSET
        '            Case swConstraintType_ALONGX
        '            Case swConstraintType_ALONGY
        '            Case swConstraintType_ALONGZ
        '            Case swConstraintType_ALONGXPOINTS
        '            Case swConstraintType_ALONGYPOINTS
        '            Case swConstraintType_ALONGZPOINTS
        '            Case swConstraintType_PARALLELYZ
        '            Case swConstraintType_PARALLELZX
		Case swConstraintType_INTERSECTION
			strType = "intersection"
			kEnt1 = Array("source", define_entity(vEnts(0), False))
			kEnt2 = Array("destination1", define_entity(vEnts(1), False))
			kEnt3 = Array("destination2", define_entity(vEnts(2), False))
        '            Case swConstraintType_PATTERNED
        '            Case swConstraintType_ISOBYPOINT
        '            Case swConstraintType_SAMEISOPARAM
        '            Case swConstraintType_FITSPLINE
		
		Case Else
			Debug.Print("Relation type " & swSketchRel.GetRelationType & " not yet supported")
	End Select
	
	' Set type field
	kType = Array("type", strType)
	explode_relation.Add kType, kType(0)
	
	' Figure out how many ents are defined
	If Not IsEmpty(kEnt3) Then
		explode_relation.Add kEnt3, kEnt3(0)
	End If
	If Not IsEmpty(kEnt2) Then
		explode_relation.Add kEnt2, kEnt2(0)
	End If
	If Not IsEmpty(kEnt1) Then
		explode_relation.Add kEnt1, kEnt1(0)
	End If

    ' Add the sw_def field
    kdef_sw = Array(con_sw, def_sw)
    explode_relation.Add kdef_sw, kdef_sw(0).Item("key")
End Function