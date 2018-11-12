Attribute VB_Name = "Test"
Sub testimage()

Dim raster As AcadLayer
Dim image As AcadRasterImage
Dim scalefactor As Double
Dim rotAngle As Double

Dim insertionPoint(0 To 2) As Double

Set raster = ThisDrawing.Layers.Add("Raster")
Application.ActiveDocument.ActiveLayer = raster

'Insert Image into AutoCAD
insertionPoint(0) = 0
insertionPoint(1) = 0
insertionPoint(2) = 0
scalefactor = 40
rotAngle = 0

Set image = Application.ActiveDocument.ModelSpace.AddRaster(PathToRasterImage, _
insertionPoint, scalefactor * 36, rotAngle)
image.Transparency = True
image.scalefactor = scalefactor
image.Name = "New Image"

Application.ActiveDocument.Regen acActiveViewport
Application.ActiveDocument.Save
End Sub
Sub Example_AddRasterExample_AddRaster()
    ' This example adds a raster image in model space.
    
    ' This example uses the "downtown.jpg" found in the sample
    ' directory. If you do not have this image, or it is located
    ' in a different directory, insert a valid path and filename
    ' for the imageName variable below.
    
    Dim insertionPoint(0 To 2) As Double
    Dim scalefactor As Double
    Dim rotationAngle As Double
    Dim imageName As String
    Dim rasterObj As AcadRasterImage
    imageName = "C:/violin01.jpg"
    insertionPoint(0) = 5#: insertionPoint(1) = 5#: insertionPoint(2) = 0#
    scalefactor = 1#
    rotationAngle = 0
    
    On Error Resume Next
    ' Creates a raster image in model space
    Set rasterObj = ThisDrawing.ModelSpace.AddRaster(imageName, insertionPoint, scalefactor, rotationAngle)
    Debug.Print imageName
    If Err.Description = "File error" Then
        MsgBox imageName & " could not be found."
        Exit Sub
    End If
End Sub

