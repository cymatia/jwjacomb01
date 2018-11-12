Attribute VB_Name = "Mod_Cad"
Public Sub InsertJBlock(Filename As String, TblPT As Variant, Space As String)
Dim XYZScale As Double
Dim Rotation As Double

 

  'InsertionPoint(0) = 0#: InsertionPoint(1) = 0#: InsertionPoint(2) = 0#
  XYZScale = 1#
  Rotation = 0#
 
  If Space = "Model" Then
    ThisDrawing.ModelSpace.InsertBlock TblPT, Filename, XYZScale, XYZScale, XYZScale, Rotation
   
   
Else
      ThisDrawing.PaperSpace.InsertBlock TblPT, Filename, XYZScale, XYZScale, XYZScale, Rotation
  End If
End Sub

Public Sub InsertJBlock1(Filename As String, TblPT As Variant, Space As String)
Dim XYZScale As Double
Dim Rotation As Double

 

  'InsertionPoint(0) = 0#: InsertionPoint(1) = 0#: InsertionPoint(2) = 0#
  XYZScale = 1#
  Rotation = 0#
 
  If Space = "Model" Then
    ThisDrawing.ModelSpace.InsertBlock TblPT, Filename, XYZScale, XYZScale, XYZScale, Rotation
   
   
Else
      ThisDrawing.PaperSpace.InsertBlock TblPT, Filename, XYZScale, XYZScale, XYZScale, Rotation
  End If
End Sub

