Attribute VB_Name = "REF_FORMAT_MTEXT"
'https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2015/ENU/AutoCAD-ActiveX/files/GUID-A43040EF-F9C5-4AEF-B87B-375FB16C77F2-htm.html
Sub Ch4_FormatMText()
  Dim mtextObj As AcadMText
  Dim insertPoint(0 To 2) As Double
  Dim width As Double
  Dim textString As String

  insertPoint(0) = 2
  insertPoint(1) = 2
  insertPoint(2) = 0
  width = 4

  ' Define the ASCII characters for the control characters
  Dim OB As Long  ' Open Bracket  {
  Dim CB As Long  ' Close Bracket }
  Dim BS As Long  ' Back Slash    \
  Dim FS As Long  ' Forward Slash /
  Dim SC As Long  ' Semicolon     ;
  OB = Asc("{")
  CB = Asc("}")
  BS = Asc("\")
  FS = Asc("/")
  SC = Asc(";")

  ' Assign the text string the following line of control
  ' characters and text characters:
  ' {{\H1.5x; Big text}\A2; over text\A1;/\A0; under text}

  textString = Chr(OB) + Chr(OB) + Chr(BS) + "H1.5x" _
  + Chr(SC) + "Big text" + Chr(CB) + Chr(BS) + "A2" _
  + Chr(SC) + "over text" + Chr(BS) + "A1" + Chr(SC) _
  + Chr(FS) + Chr(BS) + "A0" + Chr(SC) + "under text" _
  + Chr(CB)

  ' Create a text Object in model space
  Set mtextObj = ThisDrawing.ModelSpace.AddMText(insertPoint, width, textString)
  ZoomAll
End Sub

Function BigText(xText, xScFBig As Variant)
Dim textString As String
  ' Define the ASCII characters for the control characters
  Dim OB As Long  ' Open Bracket  {
  Dim CB As Long  ' Close Bracket }
  Dim BS As Long  ' Back Slash    \
  Dim FS As Long  ' Forward Slash /
  Dim SC As Long  ' Semicolon     ;
  OB = Asc("{")
  CB = Asc("}")
  BS = Asc("\")
  FS = Asc("/")
  SC = Asc(";")

  ' Assign the text string the following line of control
  ' characters and text characters:
  ' {{\H1.5x; Big text}\A2; over text\A1;/\A0; under text}
'BigText = Chr(OB) + Chr(OB) + Chr(BS) + "H1.2x" _
  + Chr(SC) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + xText + Chr(CB) + Chr(BS) + "A" _
  + Chr(SC) _
  + Chr(CB)
  'xScFBig not used
  BigText = Chr(OB) + Chr(OB) + Chr(BS) + "H" & 1.2 & "x" _
  + Chr(SC) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + xText + Chr(CB) + Chr(BS) + "A" _
  + Chr(SC) _
  + Chr(CB)

End Function
