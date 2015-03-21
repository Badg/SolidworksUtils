
Dim swApp               As SldWorks.SldWorks
Dim swModel             As SldWorks.ModelDoc2
Dim featureMgr          As SldWorks.FeatureManager
Dim swSketchMgr         As SldWorks.SketchManager
Dim swSkRelMgr          As SldWorks.SketchRelationManager
Dim modelName           As String
Dim logFileName         As String
Dim fileNumber          As Integer
Dim indent              As String
Dim bullet              As String



Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()
    'http://help.solidworks.com/2013/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeatureManager_members.html

    'Some general setups (See help)
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc

    '!!!We always starting from what we know exactly!

    'This selects the Top Plane
    boolstatus = Part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)

    '!!!Then everything we are creating we naming by the name we want, so later we can select it!

    'This creates a new plane, and we know exactly it's name, so we can refer to it later
    Dim Hardi_Plane As Object
    Set Hardi_Plane = Part.FeatureManager.InsertRefPlane(8, 0, 0, 0, 0, 0)

    ' Part.ClearSelection2 True 'unselect everything

    'We need to create the scetch on the plane that we've just created
    'So let's select our plane which is Hardi_Plane
    ' Set boolstatus = Part.Extension.SelectByID2(Hardi_Plane.Name, "PLANE", 0, 0, 0, True, 0, Nothing, 0)
    'And insert the scetch
    Part.SketchManager.InsertSketch True
    'Draw some rectangle
    Dim vSkLines As Variant
    vSkLines = Part.SketchManager.CreateCornerRectangle(-1, -1, 0, 1, 1, 0)
    '!!!And do not forget to remember the new created sketch as we need to refer to it later!
    Dim Sketch_Hardi As Object
    Set Sketch_Hardi = Part.GetActiveSketch2
    'Finilise everything
    Part.SketchManager.InsertSketch True

    Part.ClearSelection2 True 'unselect everything

    'OK, now we know all the names we want, so let's extrude something out of the sketch we've just created
    'So we select this sketch
    boolstatus = Part.Extension.SelectByID2(Sketch_Hardi.Name, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
    'And extrude it
    Dim Hardi_Extrude As Object
    Set Hardi_Extrude = Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, 0.05, 0.01, False, False, False, False, 0.01745329251994, 0.01745329251994, False, False, False, False, True, True, True, 0, 0, False)

    'That's it, Enjoy
End Sub
