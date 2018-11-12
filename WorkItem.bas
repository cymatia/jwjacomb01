Attribute VB_Name = "WorkItem"
Public Sub ShowWorkItem()
frmWorkItem.show
End Sub
Public Sub writeAssWorkItemImg()


Dim DB As Database
Dim rstProjDiv, rstProjAss, rstAssWorkItem, rstWorkImg As Recordset
Dim strProjDiv, strProjAss, strAssWorkItem, strWorkImg, txtRow As String
Dim rnProjDiv, rnProjAss, rnAssWorkItem, rnWorkImg As Integer     'ProjSet no.
'Dim rcArray As Variant
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
Dim tmpArrayAss(0 To 1) As Double

Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
Dim height As Double
Dim mode As Long
Dim prompt As String

'**********IMG********************
Dim nHeadDwg, nDivDwg, nWkCsiDwg, nRowDwg, nFootDwg, nImgDwg As Variant
Dim nJumpHead, nJumpGroup, nJumpRow As Integer
'**********IMG********************
Dim blockRefObj As AcadBlockReference
Dim BlkObj As Object
Dim xDivNo, xAssID, i, iDiv, irnDiv, irnAss, j, k, irnWorkItem, iattr_DIV, iattr_ASS, iattr_WORKITEM As Integer
Dim mg, iattr_Img As Integer

'frmWorkItem.ListBox1.Clear
CreateCNstyle
TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify

''InsertJBlock "e:\drawfile\jvba\XHDWRHEAD01.dwg", tblPT, "Model"
'ThisDrawing.ModelSpace.InsertBlock tblPT, "e:\drawfile\jvba\XHDWRHEAD02.dwg", 1, 1, 1, 0
'InsertJBlock "e:\drawfile\jvba\XHDWRGPROW02.dwg", tblPT2, "Model"
'InsertJBlock "e:\drawfile\jvba\XHDWRROW02.dwg", tblPT2, "Model"
'InsertJBlock "e:\drawfile\jvba\XHDWRFOOT02.dwg", tblPT2, "Model"

'ThisDrawing.SelectionSets.Item("XHDWRGPROW01").Delete
'ThisDrawing.SelectionSets.Item("XHDWRROW01").Delete

If frmWorkItem.ShowImage = True Then
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_HEAD01.dwg"
    'nGroupDwg = "H:\drawfile\jvba\XHDWRGPROW03.dwg"
    nDivDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_DIV_ROW01.dwg"
    nWkCsiDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_WKCSI_ROW01.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_ROW01.dwg"
    nImgFillDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_ImgFill01.dwg"
    nFootDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_FOOT01.dwg"
    nJumpHead = 60 'was 24
    nJumpDiv = 20 'was 66
    nJumpWkCSI = 20
    nJumpRow = 72 'was 60
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_HEAD10.dwg", TblPT, "Model"
Else
    'nHeadDwg = "H:\drawfile\jvba\X_WORKITEM_HEAD02.dwg"
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_HEAD10.dwg"
    'nGroupDwg = "H:\drawfile\jvba\XHDWRGPROW03.dwg"
    nDivDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_DIV_ROW02.dwg"
    nWkCsiDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_WKCSI_ROW01.dwg"
    'nRowDwg = "H:\drawfile\jvba\X_WORKITEM_ROW02.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_ROW10.dwg"
    'nImgFillDwg = "H:\drawfile\jvba\X_WORKITEM_ImgFill03.dwg" 'Why is this 03?
    nImgFillDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_ImgFill01.dwg"
    nFootDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_FOOT01.dwg"
    nJumpHead = 60 'was 24
    nJumpDiv = 20 'was 66
    nJumpWkCSI = 20
    nJumpRow = 72 'was 60
    'InsertJBlock "H:\drawfile\jvba\X_WORKITEM_HEAD02.dwg", tblPT, "Model"
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_HEAD10.dwg", TblPT, "Model"
End If

TblPT(1) = TblPT(1) - nJumpHead

'Set DB = OpenDatabase("H:\db\db_est\H_H_EST_2013_001.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_est\x_EST_2013_001.mdb")

Dim varItm As Variant


With frmWorkItem.ListBox1
For il = 0 To .ListCount - 1
If .Selected(il) Then




    'xAssDiv = !assdiv
    'xAssID = ListBox1.List(il, 0)
'    Debug.Print .Column(il, 0)
    Debug.Print .Column(0, il)
    xAssID = .Column(0, il)
    Debug.Print xAssID
    'strProjAss = "SELECT ASS.ASSID, ASS.ESTID, ASS.ASSDIV, ASS.ASSNO, ASS.ASSSUBDIV, ass.csi, ASS.ASSNAME, ASS.proj_no " _
    & "FROM ASS " _
    & "WHERE (((ASS.proj_no) = '" & frmWorkItem.seleproj & "'))" _
    & "ORDER BY ASS.ASSDIV, ASS.ASSNO"
    'strProjAss = "SELECT ASS.ASSID, ASS.ESTID, ASS.ASSDIV, ASS.ASSNO, ASS.ASSSUBDIV, ASS.CSI, ASS.ASSNAME, ASS.proj_no " _
    & "FROM ASS " _
    & "WHERE (((ASS.ASSID) =" & xAssID & "))" _
    & "ORDER BY ASS.CSI;"
If frmWorkItem.mCategory = True Then
    strProjAss = "SELECT ASS.ASSID, ASS.ESTID, ASS.ASSDIV, ASS.ASSSUBDIV, ASS.ASSNO, ASS.ASSNAME, ASS.ASSNOTE, ASS.ASSDATE, " _
    & "ASS.proj_no, ASS.WKID, ASS.CSID, ASS.CSI, ASS.ASSAMT, ASS.BUAMT, ASS.MASSAMT, " _
    & "ASS.CATEGORY_ID, CATEGORY.category_no, CATEGORY.category, CATEGORY.category_note " _
    & "FROM ASS INNER JOIN CATEGORY ON ASS.CATEGORY_ID=CATEGORY.category_id " _
    & "WHERE (((ASS.ASSID) =" & xAssID & "))" _
    & "ORDER BY CATEGORY.category_no;"
Else
    strProjAss = "SELECT ASS.ASSID, ASS.ESTID, ASS.ASSDIV, ASS.ASSSUBDIV, ASS.ASSNO, ASS.ASSNAME, ASS.ASSNOTE, ASS.ASSDATE, " _
    & "ASS.proj_no, ASS.WKID, ASS.CSID, ASS.CSI, ASS.ASSAMT, ASS.BUAMT, ASS.MASSAMT, " _
    & "ASS.CATEGORY_ID, CATEGORY.category_no, CATEGORY.category, CATEGORY.category_note " _
    & "FROM ASS INNER JOIN CATEGORY ON ASS.CATEGORY_ID=CATEGORY.category_id " _
    & "WHERE (((ASS.ASSID) =" & xAssID & "))" _
    & "ORDER BY ASS.CSI;"
End If
    
    Set rstProjAss = DB.OpenRecordset(strProjAss)
    Debug.Print rstProjAss.RecordCount
        If rstProjAss.RecordCount = 0 Then
            'GoTo Exit_writeAssWorkItemImg
        End If
    Debug.Print rstProjAss.RecordCount
    rstProjAss.MoveLast
    rstProjAss.MoveFirst
    Debug.Print rstProjAss.RecordCount
    Debug.Print rnProjAss
    rnProjAss = rstProjAss.RecordCount
        
        With rstProjAss
            Debug.Print !ASSID & "<<<<assid"
            Debug.Print !csi & "<<<<CSI"
            Debug.Print !ASSNAME & "<<<<ASSNAME"
            'xAssID = !assid
            'Debug.Print xAssID
            For j = 1 To rnProjAss
                xAssID = !ASSID
                Debug.Print xAssID
                Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                             (TblPT, nWkCsiDwg, 1, 1, 1, 0)
                Dim varAttributes2 As Variant
                varAttributes2 = blockRefObj.GetAttributes
                For iattr_ASS = LBound(varAttributes2) To UBound(varAttributes2)
                    Debug.Print varAttributes2(iattr_ASS).TagString
                    If varAttributes2(iattr_ASS).TagString = "XASSCSI" Then
                        If IsNull(!csi) = True Then
                            varAttributes2(iattr_ASS).textString = ""
                        Else
                        '******** Edited 8/14/16************
                            If frmWorkItem.mCategory = True Then
                            varAttributes2(iattr_ASS).textString = !category_no
                            Debug.Print !category_no
                            Else
                            varAttributes2(iattr_ASS).textString = !csi
                            End If
                        '***********************************
                            Debug.Print varAttributes2(iattr_ASS).textString
                        End If
                    End If
                    If varAttributes2(iattr_ASS).TagString = "XASSNAME" Then
                        Debug.Print !ASSNAME & "<<<< ASSNAME"
                        If IsNull(!ASSNAME) = True Then
                            varAttributes2(iattr_ASS).textString = ""
                        Else
                            '=DLookUp("csiname","csi","left(csi,2) = " & CStr([assdiv]))
                        '******** Edited 8/14/16************
                            If frmWorkItem.mCategory = True Then
                            varAttributes2(iattr_ASS).textString = !category
                            Else
                            varAttributes2(iattr_ASS).textString = !ASSNAME
                            End If
                        '***********************************
                            Debug.Print varAttributes2(iattr_ASS).textString
                        End If
                    End If
                Next iattr_ASS
                TblPT(1) = TblPT(1) - nJumpWkCSI
                tempYpos = TblPT(1)
                strAssWorkItem = "SELECT WORK.ASSID, WORK.WKID, WORK.WKCSINO, WORK.WKCSINAME, WORK.WKNOTE ,WORK.WKNOTE_EX " _
                & "FROM ASS INNER JOIN [WORK] ON ASS.ASSID = WORK.ASSID " _
                & "WHERE (((WORK.ASSID) =" & xAssID & "))" _
                & "ORDER BY WORK.WKCSINO;"
                'Debug.Print xwkID
                Set rstAssWorkItem = DB.OpenRecordset(strAssWorkItem)
                Debug.Print rstAssWorkItem.RecordCount
                'Debug.Print strHDSet
                If rstAssWorkItem.RecordCount = 0 Then
                    GoTo Exit_writeAssWorkItemImg
                End If
                rstAssWorkItem.MoveLast
                rstAssWorkItem.MoveFirst
                Debug.Print rstAssWorkItem.RecordCount
                Debug.Print rnAssWorkItem
                rnAssWorkItem = rstAssWorkItem.RecordCount
                With rstAssWorkItem
                    'xwkid = !wkid
                    Debug.Print xwkid
                    For k = 1 To rnAssWorkItem
                        xwkid = !wkid
                        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                     (TblPT, nRowDwg, 1, 1, 1, 0)
    '**************************************************************************************
                        Dim imgPT(0 To 2) As Double
                        Dim fillPT(0 To 2) As Double
                        Dim scalefactor, xSF As Double
                        Dim rotAngleInDegree As Double, rotAngle As Double
                        'Dim imageName As String
                        Dim imageName As String
                        Dim raster As AcadRasterImage
                        Dim imgheight As Variant
                        Dim imgwidth As Variant
                        Dim width As Double
                        Dim corner(0 To 2) As Double
                        Dim corner2(0 To 2) As Double
                        strWorkImg = "SELECT work.wkid, connwkpic.picid, PROJ_PIC.PICFULLNAME, PROJ_PIC.PICFULLNAMESM " _
                                & "FROM [WORK] INNER JOIN (connwkpic INNER JOIN PROJ_PIC ON connwkpic.picid = PROJ_PIC.PICID) " _
                                & "ON WORK.WKID = connwkpic.workid " _
                                & "WHERE (((work.wkid)=" & xwkid & "));"
    
                        Set rstWorkImg = DB.OpenRecordset(strWorkImg)
                        rnWorkImg = rstWorkImg.RecordCount
                        'Do While Not rstWorkImg.EOF
                            Debug.Print rstWorkImg.RecordCount
                            If rstWorkImg.RecordCount = 0 Then
                                GoTo SkipImg
                            Else
                                rstWorkImg.MoveLast
                                rstWorkImg.MoveFirst
                            End If
                            Debug.Print rstWorkImg.RecordCount
                            Debug.Print rnWorkImg
                            rnWorkImg = rstWorkImg.RecordCount
                            nImgFill = Round(rnWorkImg / 2, 0)
                            Debug.Print nImgFill & "<<<<<< nImgFill >>>>"
                                imgPT(0) = TblPT(0) + 6
                                imgPT(1) = TblPT(1) - 68
                                imgPT(2) = TblPT(2)
        
                            With rstWorkImg 'img1
                                rstWorkImg.MoveLast
                                rstWorkImg.MoveFirst
                                'If rstWorkImg.RecordCount = 0 Then
                                   ' Exit For
                                'End
                                For mg = 1 To rnWorkImg
                                    'If !PICFULLNAME <> "" Then
                                    If mg = 2 Then
                                        imgPT(0) = TblPT(0) + 66
                                    End If
                                    If frmWorkItem.ShowImage = True Then
                                
                                    'imageName = "C:\AutoCAD\sample\downtown.jpg"
                                    If IsNull(!picfullname) = False Then
                                        If IsNull(!picfullnamesm) = False Then
                                        imageName = !picfullnamesm
                                        Else
                                        imageName = !picfullname
                                        End If
                                    Else
                                        imageName = ""
                                    End If
                                    Debug.Print mg & ": " & imageName
                                    corner(0) = imgPT(0) + 12
                                    corner(1) = imgPT(1) - 1
                                    corner(2) = imgPT(2)
                                    width = 48 '4' X 12
                                    Set mtextObj = ThisDrawing.ModelSpace.AddMText(corner, width, !picid)
                                        'MTextObj.StyleName = "R"
                                        'MTextObj.Layer = "DI48"
                                        mtextObj.Color = acRed
                                        mtextObj.height = 2
                                        mtextObj.Update
                                        corner2(0) = imgPT(0) + 12
                                        corner2(1) = imgPT(1) + 12
                                        corner2(2) = imgPT(2)
                                    Set mtextObj = ThisDrawing.ModelSpace.AddMText(corner2, width, mg)
                                        'CreateARstyle
                                        'MTextObj.StyleName = "R"
                                        'MTextObj.Layer = "DEFPOINTS"
                                        mtextObj.height = 2
                                        mtextObj.Update
                                        'Debug.Print imageName
                                        'insertionPoint(0) = 2#: insertionPoint(1) = 2#: insertionPoint(2) = 0#
                                        scalefactor = 48#
                                        'rotAngleInDegree = 0#
                                        'rotAngle = rotAngleInDegree * 3.141592 / 180#
                                        rotAngle = 0#
                                        ' Creates a raster image in model space
                                        'If !fximgname <> "" Then
                                        Debug.Print imageName
                                        Debug.Print imgPT(0) & ":" & imgPT(1) & ":" & imgPT(2)
                                        Debug.Print scalefactor
                                        Debug.Print rotAngle
                                        
                                    Set raster = ThisDrawing.ModelSpace.AddRaster(imageName, imgPT, scalefactor, rotAngle)
                                    If Err.Description = "File error" Then
                                        MsgBox imageName & " could not be found."
                                        Exit Sub
                                    End If
                                        With raster
                                            imgheight = raster.ImageHeight
                                            imgwidth = raster.ImageWidth
                                                If imgheight > 48 Then
                                                    xSF = (48 / imgheight) * 48
                                                    Debug.Print "xSF: " & xSF * 48
                                                     raster.scalefactor = xSF
                                                End If
                                        End With
                                     End If
                                    'End If
                                    'imgPT(1) = imgPT(1) - nJumpRow
                                    If ((mg Mod 2) = 0) Then 'even
                                        Debug.Print mg Mod 2 & "see mod"
                                        imgPT(0) = TblPT(0) + 66
                                        imgPT(1) = imgPT(1) - 56
                                    Else
                                        imgPT(0) = TblPT(0) + 6
                                        'imgPT(1) = imgPT(1) - 54
                                    End If
                                    Debug.Print "mg:>>>> " & mg & " of Total: " & rnWorkImg & "Image: " & imageName
                                    .MoveNext
                                Next mg
                                '.MoveNext
                            End With 'img1
                            '.MoveNext
                        'Loop
    '**************************************************************************************
                        Dim CsiPos(0 To 2) As Double
                        CsiPos(0) = TblPT(0) - 360
                        CsiPos(1) = TblPT(1)
                        CsiPos(2) = TblPT(2)
                        width = 240 '4' X 12
                        CSILabel = UCase(!wkcsino & "-" & !WKCSINAME)
                        CreateRstyle
                        Set mtextObj = ThisDrawing.ModelSpace.AddMText(CsiPos, width, CSILabel)
                            mtextObj.StyleName = "R"
                            mtextObj.height = 4
                            mtextObj.Layer = "DI48"
                            mtextObj.Update
                        CsiPos(0) = CsiPos(0) - 240
                        width = 84 '4' X 12
                        Set mtextObj = ThisDrawing.ModelSpace.AddMText(CsiPos, width, !wkcsino)
                            mtextObj.StyleName = "R"
                            mtextObj.height = 4
                            mtextObj.Layer = "DI48"
                            mtextObj.Update
    '**************************************************************************************
SkipImg:
                        Dim varAttributes3 As Variant
                        varAttributes3 = blockRefObj.GetAttributes
    
                        For iattr_WKCSI = LBound(varAttributes3) To UBound(varAttributes3)
                            Debug.Print varAttributes3(iattr_WKCSI).TagString
                            If varAttributes3(iattr_WKCSI).TagString = "XWKCSINO" Then
                                If IsNull(!wkcsino) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = !wkcsino
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XWKCSINAME" Then
                                If IsNull(!WKCSINAME) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = !WKCSINAME
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNOTE1" Then
                                If IsNull(!WKNOTE_EX) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 1, 5, 22)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNOTE2" Then
                                If IsNull(!WKNOTE_EX) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 2, 5, 22)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNOTE2" Then
                                If IsNull(!WKNOTE_EX) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 3, 5, 22)
                                End If
                            End If
                                     If varAttributes3(iattr_WKCSI).TagString = "XNOTE3" Then
                                If IsNull(!WKNOTE_EX) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 4, 5, 22)
                                End If
                            End If
                              If varAttributes3(iattr_WKCSI).TagString = "XNOTE4" Then
                                If IsNull(!WKNOTE_EX) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 5, 5, 22)
                                End If
                            End If
                              If varAttributes3(iattr_WKCSI).TagString = "XNOTE5" Then
                                If IsNull(!WKNOTE_EX) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 5, 5, 22)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE1" Then
                                If IsNull(!WKNOTE) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 1, 5, 22)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE2" Then
                                If IsNull(!WKNOTE) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 2, 5, 22)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE3" Then
                                If IsNull(!WKNOTE) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 3, 5, 22)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE4" Then
                                If IsNull(!WKNOTE) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 4, 5, 22)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE5" Then
                                If IsNull(!WKNOTE) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 5, 5, 22)
                                End If
                            End If
    '**************************************************************************************
                            If varAttributes3(iattr_WKCSI).TagString = "XBY1" Then
                                Debug.Print WRITEWORKBY(!wkid)
                                If IsNull(WRITEWORKBY(!wkid)) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(WRITEWORKBY(!wkid), 1, 2, 22)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XBY2" Then
                                If IsNull(WRITEWORKBY(!wkid)) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(WRITEWORKBY(!wkid), 2, 2, 22)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XACT1" Then
                                If IsNull(printAction(!wkid)) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(printAction(!wkid), 1, 2, 22)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XACT2" Then
                                If IsNull(printAction(!wkid)) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(printAction(!wkid), 2, 2, 22)
                                End If
                            End If

    '**************************************************************************************
                        Next iattr_WKCSI
                        
    '******* FILL DWG ************** GOT IT TO WORK ON 11/11/12 *****************************
                        nImgFill = Round(rnWorkImg / 2, 0)
                        Debug.Print "k>>>>>>>>>>>>>>: " & k
                        Debug.Print TblPT(0) & " : " & TblPT(1) & " : " & TblPT(2) & "<<<<<<< tblPT"
                        fillPT(0) = TblPT(0)
                        fillPT(1) = TblPT(1) - nJumpRow
                        fillPT(2) = TblPT(2)
                        Debug.Print fillPT(0) & " : " & fillPT(1) & " : " & fillPT(2) & "<<<<<< fillPT"
                        Debug.Print tempYpos
                        If nImgFill >= 2 Then
                            For iFill = 1 To nImgFill - 1
                                Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                    (fillPT, nImgFillDwg, 1, 1, 1, 0)
                                fillPT(1) = fillPT(1) - 56
                            Next iFill
                        End If
                            If nImgFill < 2 Then
                                TblPT(1) = TblPT(1) - nJumpRow
                            Else
                                TblPT(1) = TblPT(1) - nJumpRow - ((nImgFill - 1) * 56)
                            End If
     '***** FILL DWG *************************************************************************
                        .MoveNext
                    Next k 'rnAssWorkItem
                    '.MoveNext
                End With
                'rstWorkItem.Close
                .MoveNext
            Next j 'rnProjAss
            
    End With 'rstProjAss first with
End If
Next il
End With
'rstProjAss.Close
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                             (TblPT, nFootDwg, 1, 1, 1, 0)
''ThisDrawing.SelectionSets.Item("hdwrsched").Delete
'rstProjDiv.Close
'rstProjAss.Close
'rstWorkItem.Close


'RsHDwrSet.Close
Set DB = Nothing

''Else

''ThisDrawing.SelectionSets.Item("hdwrsched").Delete
'no attribute block, inform the user
'

'MsgBox "Inserted block! Re-run this program.", vbCritical, "JWJA"

'delete the selection set
'ThisDrawing.SelectionSets.Item("TBLK").Delete

''End If

'Set MTextObj = ThisDrawing.ModelSpace.AddMText(corner, width, text)

''ThisDrawing.SelectionSets.Item("hdwrsched").Delete


Exit_writeAssWorkItemImg:
frmWorkItem.Hide

End Sub
Public Sub writeAssWorkItemImgCAD()
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+          11/14/12 FOR CAD USE WITH NO OF IMAGE CONTROL
'+          03/18/14 FOR LARGE PICTURE
'+
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Dim DB As Database
Dim rstProjDiv, rstProjAss, rstAssWorkItem, rstWorkImg As Recordset
Dim strProjDiv, strProjAss, strAssWorkItem, strWorkImg, txtRow As String
Dim rnProjDiv, rnProjAss, rnAssWorkItem, rnWorkImg As Integer     'ProjSet no.
'Dim rcArray As Variant
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
Dim tmpArrayAss(0 To 1) As Double

Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
Dim height As Double
Dim mode As Long
Dim prompt As String

'**********IMG********************
Dim nHeadDwg, nDivDwg, nWkCsiDwg, nRowDwg, nFootDwg, nImgDwg As Variant
Dim nJumpHead, nJumpGroup, nJumpRow As Integer
'**********IMG********************
Dim blockRefObj As AcadBlockReference
Dim BlkObj As Object
Dim xDivNo, xAssID, i, iDiv, irnDiv, irnAss, j, k, irnWorkItem, iattr_DIV, iattr_ASS, iattr_WORKITEM As Integer
Dim mg, iattr_Img As Integer

'frmWorkItem.ListBox1.Clear
'CreateCNstyle
TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify


If frmWorkItem.ShowImage = True Then
    If frmWorkItem.PrintLarge = True Then
        nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XCSI_HD_10.dwg"
        'nGroupDwg = "H:\drawfile\jvba\XHDWRGPROW03.dwg"
        nDivDwg = "\\jwja-svr-10\drawfile\jvba\XCSI_DIV_01.dwg"
        nWkCsiDwg = "\\jwja-svr-10\drawfile\jvba\XCSI_ASS_01.dwg"
        nRowDwg = "\\jwja-svr-10\drawfile\jvba\XCSI_WORK_ROW_10.dwg"
        'nImgFillDwg = "H:\drawfile\jvba\X_WORKITEM_ImgFill01.dwg"
        nFootDwg = "\\jwja-svr-10\drawfile\jvba\XCSI_FOOT_01.dwg"
        nJumpHead = 42 'was 60
        nJumpDiv = 18 'was 20
        nJumpWkCSI = 18 'WAS 20
        nJumpRow = 132 'was 72
        InsertJBlock "\\jwja-svr-10\drawfile\jvba\XCSI_HD_10.dwg", TblPT, "Model"

    Else
        nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XCSI_HD_01.dwg"
        'nGroupDwg = "H:\drawfile\jvba\XHDWRGPROW03.dwg"
        nDivDwg = "\\jwja-svr-10\drawfile\jvba\XCSI_DIV_01.dwg"
        nWkCsiDwg = "\\jwja-svr-10\drawfile\jvba\XCSI_ASS_01.dwg"
        nRowDwg = "\\jwja-svr-10\drawfile\jvba\XCSI_WORK_ROW_01.dwg"
        'nImgFillDwg = "H:\drawfile\jvba\X_WORKITEM_ImgFill01.dwg"
        nFootDwg = "\\jwja-svr-10\drawfile\jvba\XCSI_FOOT_01.dwg"
        nJumpHead = 42 'was 60
        nJumpDiv = 18 'was 20
        nJumpWkCSI = 18 'WAS 20
        nJumpRow = 72 'was 72
        InsertJBlock "\\jwja-svr-10\drawfile\jvba\XCSI_HD_01.dwg", TblPT, "Model"

    End If
    
    
    'InsertJBlock "H:\drawfile\jvba\XCSI_HD_01.dwg", tblPT, "Model"
Else
    Exit Sub
End If

TblPT(1) = TblPT(1) - nJumpHead

'Set DB = OpenDatabase("H:\db\db_est\H_H_EST_2013_001.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_est\x_EST_2013_001.mdb")

Dim varItm As Variant


With frmWorkItem.ListBox1
For il = 0 To .ListCount - 1
If .Selected(il) Then




    'xAssDiv = !assdiv
    'xAssID = ListBox1.List(il, 0)
    Debug.Print .Column(il, 0)
    Debug.Print .Column(0, il)
    xAssID = .Column(0, il)
    Debug.Print xAssID
    'strProjAss = "SELECT ASS.ASSID, ASS.ESTID, ASS.ASSDIV, ASS.ASSNO, ASS.ASSSUBDIV, ass.csi, ASS.ASSNAME, ASS.proj_no " _
    & "FROM ASS " _
    & "WHERE (((ASS.proj_no) = '" & frmWorkItem.seleproj & "'))" _
    & "ORDER BY ASS.ASSDIV, ASS.ASSNO"
    strProjAss = "SELECT ASS.ASSID, ASS.ESTID, ASS.ASSDIV, ASS.ASSNO, ASS.ASSSUBDIV, ASS.CSI, ASS.ASSNAME, ASS.proj_no " _
    & "FROM ASS " _
    & "WHERE (((ASS.ASSID) =" & xAssID & "))" _
    & "ORDER BY ASS.CSI;"

    
    
    Set rstProjAss = DB.OpenRecordset(strProjAss)
    Debug.Print rstProjAss.RecordCount
        If rstProjAss.RecordCount = 0 Then
            GoTo Exit_writeAssWorkItemImgCAD
        End If
    rstProjAss.MoveLast
    rstProjAss.MoveFirst
    Debug.Print rstProjAss.RecordCount
    Debug.Print rnProjAss
    rnProjAss = rstProjAss.RecordCount
        
        With rstProjAss
            Debug.Print !ASSID & "<<<<assid"
            Debug.Print !csi & "<<<<CSI"
            Debug.Print !ASSNAME & "<<<<ASSNAME"
            'xAssID = !assid
            'Debug.Print xAssID
            For j = 1 To rnProjAss
                xAssID = !ASSID
                Debug.Print xAssID
                Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                             (TblPT, nWkCsiDwg, 1, 1, 1, 0)
                Dim varAttributes2 As Variant
                varAttributes2 = blockRefObj.GetAttributes
                For iattr_ASS = LBound(varAttributes2) To UBound(varAttributes2)
                    Debug.Print varAttributes2(iattr_ASS).TagString
                    If varAttributes2(iattr_ASS).TagString = "XASSCSI" Then
                        If IsNull(!csi) = True Then
                            varAttributes2(iattr_ASS).textString = ""
                        Else
                            varAttributes2(iattr_ASS).textString = !csi
                            Debug.Print varAttributes2(iattr_ASS).textString
                        End If
                    End If
                    If varAttributes2(iattr_ASS).TagString = "XASSNAME" Then
                        Debug.Print !ASSNAME & "<<<< ASSNAME"
                        If IsNull(!ASSNAME) = True Then
                            varAttributes2(iattr_ASS).textString = ""
                        Else
                            '=DLookUp("csiname","csi","left(csi,2) = " & CStr([assdiv]))
                            varAttributes2(iattr_ASS).textString = !ASSNAME
                            Debug.Print varAttributes2(iattr_ASS).textString
    
                        End If
                    End If
                Next iattr_ASS
                TblPT(1) = TblPT(1) - nJumpWkCSI
                tempYpos = TblPT(1)
                strAssWorkItem = "SELECT WORK.ASSID, WORK.WKID, WORK.WKCSINO, WORK.WKCSINAME, WORK.WKNOTE ,WORK.WKNOTE_EX " _
                & "FROM ASS INNER JOIN [WORK] ON ASS.ASSID = WORK.ASSID " _
                & "WHERE (((WORK.ASSID) =" & xAssID & "))" _
                & "ORDER BY WORK.WKCSINO;"
    
                'Debug.Print xwkID
                Set rstAssWorkItem = DB.OpenRecordset(strAssWorkItem)
                Debug.Print rstAssWorkItem.RecordCount
                'Debug.Print strHDSet
                If rstAssWorkItem.RecordCount = 0 Then
                    GoTo Exit_writeAssWorkItemImgCAD
                End If
                rstAssWorkItem.MoveLast
                rstAssWorkItem.MoveFirst
                Debug.Print rstAssWorkItem.RecordCount
                Debug.Print rnAssWorkItem
                
                rnAssWorkItem = rstAssWorkItem.RecordCount
                With rstAssWorkItem
                    'xwkid = !wkid
                    rnAssWorkItem = rstAssWorkItem.RecordCount
                    Debug.Print xwkid
                    For k = 1 To rnAssWorkItem
                        xwkid = !wkid
                        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                     (TblPT, nRowDwg, 1, 1, 1, 0)
                        Debug.Print "k >>>>: " & k
    '**************************************************************************************
                        Dim imgPT(0 To 2) As Double
                        Dim fillPT(0 To 2) As Double
                        
                        Dim scalefactor, xSF As Double
                        Dim rotAngleInDegree As Double, rotAngle As Double
                        'Dim imageName As String
                        Dim imageName As String
                        Dim raster As AcadRasterImage
                        Dim imgheight As Variant
                        Dim imgwidth As Variant
                        Dim width As Double
                        Dim corner(0 To 2) As Double
                        Dim corner2(0 To 2) As Double
                        Dim prtImgNo, ImgTabx, ImgTaby As Integer
    '#########################################################################################
                        prtImgNo = 2
    '#########################################################################################
                        strWorkImg = "SELECT work.wkid, connwkpic.picid, PROJ_PIC.PICFULLNAME " _
                                & "FROM [WORK] INNER JOIN (connwkpic INNER JOIN PROJ_PIC ON connwkpic.picid = PROJ_PIC.PICID) " _
                                & "ON WORK.WKID = connwkpic.workid " _
                                & "WHERE (((work.wkid)=" & xwkid & "));"
    
                        Set rstWorkImg = DB.OpenRecordset(strWorkImg)
                        If rstWorkImg.RecordCount = 0 Then
                        GoTo NoImg
                        Else
                        rstWorkImg.MoveLast
                        rstWorkImg.MoveFirst
                        End If
                        'rstWorkImg.MoveLast
                        'rstWorkImg.MoveFirst

                        Debug.Print rstWorkImg.RecordCount
                        Debug.Print rnWorkImg
                        rnWorkImg = rstWorkImg.RecordCount
                        nImgFill = Round(rnWorkImg / 2, 0)
                        nImgFill = Round((rnWorkImg - (rnWorkImg - prtImgNo)) / 2, 0)
                        mgTop = rnWorkImg - (rnWorkImg - prtImgNo)
                        Debug.Print nImgFill & "<<<<<< nImgFill >>>>"
                        If frmWorkItem.PrintLarge = True Then
                            imgPT(0) = TblPT(0) + 6
                            imgPT(1) = TblPT(1) - 128 ' was 128
                            imgPT(2) = TblPT(2)
                        Else
                            imgPT(0) = TblPT(0) + 6
                            imgPT(1) = TblPT(1) - 68
                            imgPT(2) = TblPT(2)
                        End If
                        
 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
      '                  Do While Not rstWorkImg.EOF
                        
                        With rstWorkImg 'img1
                        'Do While Not rstWorkImg.EOF
                            rstWorkImg.MoveLast
                            rstWorkImg.MoveFirst
                            Debug.Print rstWorkImg.RecordCount
                            Debug.Print rnWorkImg
                            rnWorkImg = rstWorkImg.RecordCount
                            nImgFill = Round(rnWorkImg / 2, 0)
                            nImgFill = Round((rnWorkImg - (rnWorkImg - prtImgNo)) / 2, 0)
                            mgTop = rnWorkImg - (rnWorkImg - prtImgNo)
                            'For mg = 1 To rnWorkImg
                            'For mg = 1 To rnWorkImg - (rnWorkImg - prtImgNo)
                            For mg = 1 To mgTop
                                'If !PICFULLNAME <> "" Then
                                Debug.Print ">>>>> mg>>>>" & mg
                                
                                If mg = 2 Then
                                    If frmWorkItem.PrintLarge = True Then
                                        imgPT(0) = TblPT(0) + 132
                                    Else
                                        imgPT(0) = TblPT(0) + 66
                                    End If
                                End If
                                If frmWorkItem.ShowImage = True Then
                                
                                    'imageName = "C:\AutoCAD\sample\downtown.jpg"
                                    If IsNull(!picfullname) = False Then
                                    imageName = !picfullname
                                    Else
                                    imageName = ""
                                    End If
                                    Debug.Print mg & ": " & imageName
                                    corner(0) = imgPT(0) + 12
                                    corner(1) = imgPT(1) - 1
                                    corner(2) = imgPT(2)
                                    width = 48 '4' X 12
                                    CreateRstyle
                                    Set mtextObj = ThisDrawing.ModelSpace.AddMText(corner, width, !picid)
                                        mtextObj.StyleName = "R"
                                        mtextObj.Layer = "DI48"
                                        mtextObj.Color = acRed
                                        mtextObj.height = 2
                                        mtextObj.Update
                                        'Debug.Print !picfullname
                                    corner2(0) = imgPT(0) + 12
                                    corner2(1) = imgPT(1) + 12
                                    corner2(2) = imgPT(2)
                                    Set mtextObj = ThisDrawing.ModelSpace.AddMText(corner2, width, mg)
                                        'CreateARstyle
                                        mtextObj.StyleName = "R"
                                        mtextObj.Layer = "DEFPOINTS"
                                        mtextObj.height = 2
                                        mtextObj.Update
                                    'Debug.Print imageName
                                    'insertionPoint(0) = 2#: insertionPoint(1) = 2#: insertionPoint(2) = 0#
                                    
                                        
                                    If frmWorkItem.PrintLarge = True Then
                                    scalefactor = 96#
                                    'rotAngleInDegree = 0#
                                    'rotAngle = rotAngleInDegree * 3.141592 / 180#
                                    rotAngle = 0#
                                    imgScale = 96
                                    ImgTabx = 108
                                    ImgTaby = 128
                                    Else
                                    scalefactor = 48#
                                    rotAngle = 0#
                                    imgScale = 48
                                    ImgTabx = 66
                                    ImgTaby = 56
                                    End If
                                    
                                    
                                    ' Creates a raster image in model space
                                    'If !fximgname <> "" Then
                                        Set raster = ThisDrawing.ModelSpace.AddRaster(imageName, imgPT, scalefactor, rotAngle)
                                        With raster
                                            imgheight = raster.ImageHeight
                                            imgwidth = raster.ImageWidth
                                                If imgheight > imgScale Then
                                                    xSF = (imgScale / imgheight) * imgScale
                                                    Debug.Print "xSF: " & xSF * imgScale
                                                     raster.scalefactor = xSF
                                                End If
                                        End With
                                    End If
                                'End If
                            
                            'imgPT(1) = imgPT(1) - nJumpRow
                            If ((mg Mod 2) = 0) Then 'even
                            
                            Debug.Print mg Mod 2 & "see mod"
                            
                                imgPT(0) = TblPT(0) + ImgTabx
                                
                                imgPT(1) = imgPT(1) - ImgTaby
                            Else
                                imgPT(0) = TblPT(0) + 6
                                'imgPT(1) = imgPT(1) - 54
                            End If
                            If rnWorkImg < 2 Then
                            'Exit Do
                            Exit For
                            End If
                            .MoveNext
                            Next mg
                            Debug.Print "test"
                            '.MoveNext
                        'Loop
                        End With 'img1
     '                   Loop
 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
NoImg:
    '**************************************************************************************
                        Dim varAttributes3 As Variant
                        varAttributes3 = blockRefObj.GetAttributes
    
                        For iattr_WKCSI = LBound(varAttributes3) To UBound(varAttributes3)
                            Debug.Print varAttributes3(iattr_WKCSI).TagString
                            If varAttributes3(iattr_WKCSI).TagString = "XWKCSINO" Then
                                If IsNull(!wkcsino) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = !wkcsino
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XWKCSINAME" Then
                                If IsNull(!WKCSINAME) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = !WKCSINAME
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNOTE1" Then
                                If IsNull(!WKNOTE_EX) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 1, 5, 46)
                                    'varAttributes3(iattr_WKCSI).TextString = !WKNOTE_EX
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNOTE2" Then
                                If IsNull(!WKNOTE_EX) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 2, 5, 46)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNOTE3" Then
                                If IsNull(!WKNOTE_EX) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 3, 5, 46)
                                End If
                            End If
                              If varAttributes3(iattr_WKCSI).TagString = "XNOTE4" Then
                                If IsNull(!WKNOTE_EX) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 4, 5, 46)
                                End If
                            End If
                              If varAttributes3(iattr_WKCSI).TagString = "XNOTE5" Then
                                If IsNull(!WKNOTE_EX) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 5, 5, 46)
                                End If
                            End If
                            '#################################################################################
                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE1" Then
                                If IsNull(!WKNOTE) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 1, 5, 46)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE2" Then
                                If IsNull(!WKNOTE) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 2, 5, 46)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE3" Then
                                If IsNull(!WKNOTE) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 3, 5, 46)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE4" Then
                                If IsNull(!WKNOTE) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 4, 5, 46)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE5" Then
                                If IsNull(SEPTEXT2(!WKNOTE, 4, 5, 60)) = True Then
                                'If Len(SEPTEXT2(!WKNOTE, 5, 5, 60)) > 1 Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 5, 5, 46)
                                End If
                            End If
    '**************************************************************************************
                            If varAttributes3(iattr_WKCSI).TagString = "XBY1" Then
                                Debug.Print WRITEWORKBY(!wkid)
                                If IsNull(WRITEWORKBY(!wkid)) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(WRITEWORKBY(!wkid), 1, 2, 20)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XBY2" Then
                                If IsNull(WRITEWORKBY(!wkid)) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(WRITEWORKBY(!wkid), 2, 2, 20)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XACT1" Then
                                If IsNull(printAction(!wkid)) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(printAction(!wkid), 1, 2, 20)
                                End If
                            End If
                            If varAttributes3(iattr_WKCSI).TagString = "XACT2" Then
                                If IsNull(printAction(!wkid)) = True Then
                                    varAttributes3(iattr_WKCSI).textString = ""
                                Else
                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(printAction(!wkid), 2, 2, 20)
                                End If
                            End If
    '**************************************************************************************
                        Dim CsiPos(0 To 2) As Double
                        CsiPos(0) = TblPT(0) - 360
                        CsiPos(1) = TblPT(1)
                        CsiPos(2) = TblPT(2)
                        width = 240 '4' X 12
                        CSILabel = UCase(!wkcsino & "-" & !WKNOTE_EX)
                        Set mtextObj = ThisDrawing.ModelSpace.AddMText(CsiPos, width, CSILabel)
                            mtextObj.StyleName = "R"
                            mtextObj.height = 4
                            mtextObj.Layer = "DI48"
                            mtextObj.Update
                        CsiPos(0) = CsiPos(0) - 240
                        width = 84 '4' X 12
                        If IsNull(!wkcsino) = False Then
                        Set mtextObj = ThisDrawing.ModelSpace.AddMText(CsiPos, width, !wkcsino)
                            mtextObj.StyleName = "R"
                            mtextObj.height = 4
                            mtextObj.Layer = "DI48"
                            mtextObj.Update
                        End If
    '**************************************************************************************
                        Next iattr_WKCSI
                        
    '******* FILL DWG ************** GOT IT TO WORK ON 11/11/12 *****************************
    
                        Debug.Print TblPT(0) & " : " & TblPT(1) & " : " & TblPT(2) & "<<<<<<< tblPT"
                        fillPT(0) = TblPT(0)
                        fillPT(1) = TblPT(1) - nJumpRow
                        fillPT(2) = TblPT(2)
                        Debug.Print fillPT(0) & " : " & fillPT(1) & " : " & fillPT(2) & "<<<<<< fillPT"
                        Debug.Print tempYpos
                            For iFill = 1 To nImgFill - 1
                                Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                    (fillPT, nImgFillDwg, 1, 1, 1, 0)
                                fillPT(1) = fillPT(1) - ImgTaby
                            Next iFill
                            If nImgFill <= 1 Then
                                TblPT(1) = TblPT(1) - nJumpRow
                            Else
                                TblPT(1) = TblPT(1) - nJumpRow - ((nImgFill - 1) * ImgTaby)
                            End If
     '***** FILL DWG *************************************************************************
                        .MoveNext
                    Next k 'rnAssWorkItem
                    '.MoveNext
                End With
                'rstWorkItem.Close
                .MoveNext
            Next j 'rnProjAss
            
    End With 'rstProjAss first with
End If
Next il
End With
rstProjAss.Close
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                             (TblPT, nFootDwg, 1, 1, 1, 0)
''ThisDrawing.SelectionSets.Item("hdwrsched").Delete
'rstProjDiv.Close
'rstProjAss.Close
'rstWorkItem.Close


'RsHDwrSet.Close
Set DB = Nothing

''Else

''ThisDrawing.SelectionSets.Item("hdwrsched").Delete
'no attribute block, inform the user
'

'MsgBox "Inserted block! Re-run this program.", vbCritical, "JWJA"

'delete the selection set
'ThisDrawing.SelectionSets.Item("TBLK").Delete

''End If

'Set MTextObj = ThisDrawing.ModelSpace.AddMText(corner, width, text)

''ThisDrawing.SelectionSets.Item("hdwrsched").Delete


Exit_writeAssWorkItemImgCAD:
frmWorkItem.Hide

End Sub

Public Sub writeWorkItemImg()


Dim DB As Database
Dim rstProjDiv, rstProjAss, rstAssWorkItem, rstWorkImg As Recordset
Dim strProjDiv, strProjAss, strAssWorkItem, strWorkImg, txtRow As String
Dim rnProjDiv, rnProjAss, rnAssWorkItem, rnWorkImg As Integer     'ProjSet no.
'Dim rcArray As Variant
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
Dim tmpArrayAss(0 To 1) As Double

Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
Dim height As Double
Dim mode As Long
Dim prompt As String

'**********IMG********************
Dim nHeadDwg, nDivDwg, nWkCsiDwg, nRowDwg, nFootDwg, nImgDwg As Variant
Dim nJumpHead, nJumpGroup, nJumpRow As Integer
'**********IMG********************

'Dim tag As String
'Dim value As String
Dim blockRefObj As AcadBlockReference
'Dim attData() As AcadObject

'Dim FilterGp(0) As Integer
'Dim FilterDt(0) As Variant
Dim BlkObj As Object
'Dim Pt1(0) As Double
'Dim Pt2(0) As Double

'Dim xHDWRSETID, i, iatHG, rnHDWR, k, iatHW As Integer
Dim xDivNo, xAssID, i, iDiv, irnDiv, irnAss, j, k, irnWorkItem, iattr_DIV, iattr_ASS, iattr_WORKITEM As Integer
Dim mg, iattr_Img As Integer



frmWorkItem.ListBox1.Clear
CreateCNstyle






'strHDSet = "SELECT HDWRCONN.HDWRSETID, " _
                        & "HDWRCONN.SETQTY, HDWR.TYPE, HDWR.CAT, HDWR.MFGR, HDWR.FIN, HDWR.HDWRNOTE, HDWR.HDIMGNAME, HDWR.HDWRDESC " _
                        & "FROM HDWR INNER JOIN HDWRCONN ON HDWR.HDID = HDWRCONN.HDID " _
                        & "WHERE (((HDWRCONN.HDWRSETID)=" & xHDWRSETID & "));"





TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify

''InsertJBlock "e:\drawfile\jvba\XHDWRHEAD01.dwg", tblPT, "Model"
'ThisDrawing.ModelSpace.InsertBlock tblPT, "e:\drawfile\jvba\XHDWRHEAD02.dwg", 1, 1, 1, 0
'InsertJBlock "e:\drawfile\jvba\XHDWRGPROW02.dwg", tblPT2, "Model"
'InsertJBlock "e:\drawfile\jvba\XHDWRROW02.dwg", tblPT2, "Model"
'InsertJBlock "e:\drawfile\jvba\XHDWRFOOT02.dwg", tblPT2, "Model"

'ThisDrawing.SelectionSets.Item("XHDWRGPROW01").Delete
'ThisDrawing.SelectionSets.Item("XHDWRROW01").Delete

If frmWorkItem.ShowImage = True Then
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_HEAD01.dwg"
    'nGroupDwg = "H:\drawfile\jvba\XHDWRGPROW03.dwg"
    nDivDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_DIV_ROW01.dwg"
    nWkCsiDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_WKCSI_ROW01.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_ROW01.dwg"
    nImgFillDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_ImgFill01.dwg"
    nFootDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_FOOT01.dwg"
    nJumpHead = 60 'was 24
    nJumpDiv = 24 'was 66
    nJumpWkCSI = 24
    nJumpRow = 72 'was 60
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_HEAD01.dwg", TblPT, "Model"
Else
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_HEAD02.dwg"
    'nGroupDwg = "H:\drawfile\jvba\XHDWRGPROW03.dwg"
    nDivDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_DIV_ROW02.dwg"
    nWkCsiDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_WKCSI_ROW02.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_ROW02.dwg"
    nImgFillDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_ImgFill03.dwg"
    nFootDwg = "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_FOOT02.dwg"
    nJumpHead = 60 'was 24
    nJumpDiv = 24 'was 66
    nJumpWkCSI = 24
    nJumpRow = 72 'was 60
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\X_WORKITEM_HEAD02.dwg", TblPT, "Model"
End If

TblPT(1) = TblPT(1) - nJumpHead

'Set DB = OpenDatabase("H:\db\db_est\H_H_EST_2013_001.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\Jdb\db_est\H_H_EST_2013_001.mdb")

strProjDiv = "SELECT ASS.ASSDIV, DIV.DIVNAME, ASS.proj_no " _
& "FROM ASS INNER JOIN DIV ON ASS.ASSDIV = DIV.DIVNO " _
& "GROUP BY ASS.ASSDIV, DIV.DIVNAME, ASS.proj_no " _
& "HAVING (((ASS.proj_no)='" & frmWorkItem.seleproj & "'))"

RecordSets:
Set rstProjDiv = DB.OpenRecordset(strProjDiv, dbOpenDynaset)

'Set rstProjAss = DB.OpenRecordset(strProjAss, dbOpenDynaset)
tempcount = rstProjDiv.RecordCount
If rstProjDiv.RecordCount = 0 Then
    GoTo Exit_writeWorkItemImg
End If

rstProjDiv.MoveLast
rstProjDiv.MoveFirst
rnProjDiv = rstProjDiv.RecordCount
Debug.Print rnProjAss

With rstProjDiv
    For i = 1 To rnProjDiv
         xDivNo = !assdiv
         Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, nDivDwg, 1, 1, 1, 0)
        '******************************
          'Get the block's attributes
                Dim varAttributes As Variant
                varAttributes = blockRefObj.GetAttributes
                    For iattr_DIV = LBound(varAttributes) To UBound(varAttributes)
                        Debug.Print varAttributes(iattr_DIV).TagString
                        If varAttributes(iattr_DIV).TagString = "XDIVNO" Then
                            If IsNull(!assdiv) = True Then
                                varAttributes(iattr_DIV).textString = ""
                            Else
                                varAttributes(iattr_DIV).textString = !assdiv
                                Debug.Print varAttributes(iattr_DIV).textString
                            End If

                        End If
                        If varAttributes(iattr_DIV).TagString = "XDIVNAME" Then
                            If IsNull(!DIVNAME) = True Then
                                varAttributes(iattr_DIV).textString = ""
                            Else
                                '=DLookUp("csiname","csi","left(csi,2) = " & CStr([assdiv]))
                                varAttributes(iattr_DIV).textString = !DIVNAME
                                Debug.Print varAttributes(iattr_DIV).textString

                            End If

                        End If
                    Next iattr_DIV
                    TblPT(1) = TblPT(1) - nJumpDiv
                    xAssDiv = !assdiv
                    strProjAss = "SELECT ASS.ASSID, ASS.ESTID, ASS.ASSDIV, ASS.ASSNO, ASS.ASSSUBDIV, ass.csi, ASS.ASSNAME, ASS.proj_no " _
                    & "FROM ASS " _
                    & "WHERE (((ASS.proj_no) = '" & frmWorkItem.seleproj & "'))" _
                    & "ORDER BY ASS.ASSDIV, ASS.ASSNO"

                    'Debug.Print xAssID
                    Set rstProjAss = DB.OpenRecordset(strProjAss)
                    Debug.Print rstProjAss.RecordCount
                    'Debug.Print xAssID
                    'Debug.Print strHDSet
                        If rstProjAss.RecordCount = 0 Then
                            GoTo Exit_writeWorkItemImg
                        End If
                        rstProjAss.MoveLast
                        rstProjAss.MoveFirst
                        Debug.Print rstProjAss.RecordCount
                        Debug.Print rnProjAss
                        rnProjAss = rstProjAss.RecordCount
                        
                        With rstProjAss
                            Debug.Print !ASSID & "<<<<assid"
                            Debug.Print !csi & "<<<<CSI"
                            Debug.Print !ASSNAME & "<<<<ASSNAME"
                            'xAssID = !assid
                            'Debug.Print xAssID
                            For j = 1 To rnProjAss
                                xAssID = !ASSID
                                Debug.Print xAssID
                                Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                             (TblPT, nWkCsiDwg, 1, 1, 1, 0)
                                Dim varAttributes2 As Variant
                                varAttributes2 = blockRefObj.GetAttributes
                                For iattr_ASS = LBound(varAttributes2) To UBound(varAttributes2)
                                    Debug.Print varAttributes2(iattr_ASS).TagString
                                    If varAttributes2(iattr_ASS).TagString = "XASSCSI" Then
                                        If IsNull(!csi) = True Then
                                            varAttributes2(iattr_ASS).textString = ""
                                        Else
                                            varAttributes2(iattr_ASS).textString = !csi
                                            Debug.Print varAttributes(iattr_ASS).textString
                                        End If
                                    End If
                                    If varAttributes2(iattr_ASS).TagString = "XASSNAME" Then
                                        Debug.Print !ASSNAME & "<<<< ASSNAME"
                                        If IsNull(!ASSNAME) = True Then
                                            varAttributes2(iattr_ASS).textString = ""
                                        Else
                                            '=DLookUp("csiname","csi","left(csi,2) = " & CStr([assdiv]))
                                            varAttributes2(iattr_ASS).textString = !ASSNAME
                                            Debug.Print varAttributes(iattr_ASS).textString

                                        End If
                                    End If
                                Next iattr_ASS
                                TblPT(1) = TblPT(1) - nJumpWkCSI
                                strAssWorkItem = "SELECT WORK.ASSID, WORK.WKID, WORK.WKCSINO, WORK.WKCSINAME, WORK.WKNOTE ,WORK.WKNOTE_EX " _
                                & "FROM ASS INNER JOIN [WORK] ON ASS.ASSID = WORK.ASSID " _
                                & "WHERE (((WORK.ASSID) =" & xAssID & "))" _
                                & "ORDER BY WORK.WKCSINO;"

                                'Debug.Print xwkID
                                Set rstAssWorkItem = DB.OpenRecordset(strAssWorkItem)
                                Debug.Print rstAssWorkItem.RecordCount
                                'Debug.Print strHDSet
                                If rstAssWorkItem.RecordCount = 0 Then
                                    GoTo Exit_writeWorkItemImg
                                End If
                                rstAssWorkItem.MoveLast
                                rstAssWorkItem.MoveFirst
                                Debug.Print rstAssWorkItem.RecordCount
                                Debug.Print rnAssWorkItem
                                
                                rnAssWorkItem = rstAssWorkItem.RecordCount
                                With rstAssWorkItem
                                    'xwkid = !wkid
                                    Debug.Print xwkid
                                    For k = 1 To rnAssWorkItem
                                        xwkid = !wkid
                                        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                     (TblPT, nRowDwg, 1, 1, 1, 0)

'**************************************************************************************
                                        Dim imgPT(0 To 2) As Double
                                        Dim fillPT(0 To 2) As Double
                                        
                                        Dim scalefactor, xSF As Double
                                        Dim rotAngleInDegree As Double, rotAngle As Double
                                        'Dim imageName As String
                                        Dim imageName As String
                                        Dim raster As AcadRasterImage
                                        Dim imgheight As Variant
                                        Dim imgwidth As Variant
                                        Dim width As Double
                                        Dim corner(0 To 2) As Double
                                        Dim corner2(0 To 2) As Double
                                        

                                        
                                        strWorkImg = "SELECT work.wkid, connwkpic.picid, PROJ_PIC.PICFULLNAME " _
                                                & "FROM [WORK] INNER JOIN (connwkpic INNER JOIN PROJ_PIC ON connwkpic.picid = PROJ_PIC.PICID) " _
                                                & "ON WORK.WKID = connwkpic.workid " _
                                                & "WHERE (((work.wkid)=" & xwkid & "));"

                                        Set rstWorkImg = DB.OpenRecordset(strWorkImg)
                                        If rstWorkImg.RecordCount = 0 Then
                                        Exit For
                                        Else
                                        rstWorkImg.MoveLast
                                        rstWorkImg.MoveFirst
                                        End If
                                        Debug.Print rstWorkImg.RecordCount
                                        Debug.Print rnWorkImg
                                        rnWorkImg = rstWorkImg.RecordCount
                                        nImgFill = Round(rnWorkImg / 2, 0)
                                        Debug.Print nImgFill & "<<<<<< nImgFill >>>>"
                                            imgPT(0) = TblPT(0) + 6
                                            imgPT(1) = TblPT(1) - 68
                                            imgPT(2) = TblPT(2)

                                        With rstWorkImg 'img1
                                            For mg = 1 To rnWorkImg
                                                'If !PICFULLNAME <> "" Then
                                                If mg = 2 Then
                                                    imgPT(0) = TblPT(0) + 66
                                                End If
                                                    If frmWorkItem.ShowImage = True Then
                                                
                                                    'imageName = "C:\AutoCAD\sample\downtown.jpg"
                                                    If IsNull(!picfullname) = False Then
                                                    imageName = !picfullname
                                                    Else
                                                    imageName = ""
                                                    End If
                                                    Debug.Print mg & ": " & imageName
                                                    corner(0) = imgPT(0) + 12
                                                    corner(1) = imgPT(1) - 4
                                                    corner(2) = imgPT(2)
                                                    width = 48 '4' X 12
                                                    Set mtextObj = ThisDrawing.ModelSpace.AddMText(corner, width, !picid)
                                                        mtextObj.StyleName = "R"
                                                        mtextObj.height = 2
                                                        mtextObj.Update
                                                    corner2(0) = imgPT(0) + 12
                                                    corner2(1) = imgPT(1) + 12
                                                    corner2(2) = imgPT(2)
                                                    Set mtextObj = ThisDrawing.ModelSpace.AddMText(corner2, width, mg)
                                                        CreateARstyle
                                                        mtextObj.StyleName = "ar"
                                                        mtextObj.height = 2
                                                        mtextObj.Update
                                                    'Debug.Print imageName
                                                    'insertionPoint(0) = 2#: insertionPoint(1) = 2#: insertionPoint(2) = 0#
                                                    
                                                        
                                                       
                                                    scalefactor = 48#
                                                    'rotAngleInDegree = 0#
                                                    'rotAngle = rotAngleInDegree * 3.141592 / 180#
                                                    rotAngle = 0#
                                                    
                                                    
                                                    
                                                    ' Creates a raster image in model space
                                                    'If !fximgname <> "" Then
                                                        Set raster = ThisDrawing.ModelSpace.AddRaster(imageName, imgPT, scalefactor, rotAngle)
                                                        With raster
                                                            imgheight = raster.ImageHeight
                                                            imgwidth = raster.ImageWidth
                                                                If imgheight > 48 Then
                                                                    xSF = (48 / imgheight) * 48
                                                                    Debug.Print "xSF: " & xSF * 48
                                                                     raster.scalefactor = xSF
                                                                End If
                                                        End With
                                                    End If
                                                'End If
                                            
                                            'imgPT(1) = imgPT(1) - nJumpRow
                                            If ((mg Mod 2) = 0) Then 'even
                                            
                                            Debug.Print mg Mod 2 & "see mod"
                                            
                                                imgPT(0) = TblPT(0) + 66
                                                imgPT(1) = imgPT(1) - 56
                                            Else
                                                imgPT(0) = TblPT(0) + 6
                                                'imgPT(1) = imgPT(1) - 54
                                            End If
                                            
                                            .MoveNext
                                            Next mg
                                            '.MoveNext
                                        End With 'img1
'**************************************************************************************
                                        Dim varAttributes3 As Variant
                                        varAttributes3 = blockRefObj.GetAttributes

                                        For iattr_WKCSI = LBound(varAttributes3) To UBound(varAttributes3)
                                            Debug.Print varAttributes3(iattr_WKCSI).TagString
                                            If varAttributes3(iattr_WKCSI).TagString = "XWKCSINO" Then
                                                If IsNull(!wkcsino) = True Then
                                                    varAttributes3(iattr_WKCSI).textString = ""
                                                Else
                                                    varAttributes3(iattr_WKCSI).textString = !wkcsino
                                                End If
                                            End If
                                            If varAttributes3(iattr_WKCSI).TagString = "XWKCSINAME" Then
                                                If IsNull(!WKCSINAME) = True Then
                                                    varAttributes3(iattr_WKCSI).textString = ""
                                                Else
                                                    varAttributes3(iattr_WKCSI).textString = !WKCSINAME
                                                End If
                                            End If
                                            If varAttributes3(iattr_WKCSI).TagString = "XNOTE1" Then
                                                If IsNull(!WKNOTE_EX) = True Then
                                                    varAttributes3(iattr_WKCSI).textString = ""
                                                Else
                                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 1, 5, 48)
                                                End If
                                            End If
                                            If varAttributes3(iattr_WKCSI).TagString = "XNOTE2" Then
                                                If IsNull(!WKNOTE_EX) = True Then
                                                    varAttributes3(iattr_WKCSI).textString = ""
                                                Else
                                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 2, 5, 48)
                                                End If
                                            End If
                                            If varAttributes3(iattr_WKCSI).TagString = "XNOTE2" Then
                                                If IsNull(!WKNOTE_EX) = True Then
                                                    varAttributes3(iattr_WKCSI).textString = ""
                                                Else
                                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 2, 5, 48)
                                                End If
                                            End If
                                                     If varAttributes3(iattr_WKCSI).TagString = "XNOTE3" Then
                                                If IsNull(!WKNOTE_EX) = True Then
                                                    varAttributes3(iattr_WKCSI).textString = ""
                                                Else
                                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 3, 5, 48)
                                                End If
                                            End If
                                              If varAttributes3(iattr_WKCSI).TagString = "XNOTE4" Then
                                                If IsNull(!WKNOTE_EX) = True Then
                                                    varAttributes3(iattr_WKCSI).textString = ""
                                                Else
                                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 4, 5, 48)
                                                End If
                                            End If
                                              If varAttributes3(iattr_WKCSI).TagString = "XNOTE5" Then
                                                If IsNull(!WKNOTE_EX) = True Then
                                                    varAttributes3(iattr_WKCSI).textString = ""
                                                Else
                                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE_EX, 5, 5, 48)
                                                End If
                                            End If
                                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE1" Then
                                                If IsNull(!WKNOTE) = True Then
                                                    varAttributes3(iattr_WKCSI).textString = ""
                                                Else
                                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 1, 5, 48)
                                                End If
                                            End If
                                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE2" Then
                                                If IsNull(!WKNOTE) = True Then
                                                    varAttributes3(iattr_WKCSI).textString = ""
                                                Else
                                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 2, 5, 48)
                                                End If
                                            End If
                                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE3" Then
                                                If IsNull(!WKNOTE) = True Then
                                                    varAttributes3(iattr_WKCSI).textString = ""
                                                Else
                                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 3, 5, 48)
                                                End If
                                            End If
                                            If varAttributes3(iattr_WKCSI).TagString = "XNWNOTE4" Then
                                                If IsNull(!WKNOTE) = True Then
                                                    varAttributes3(iattr_WKCSI).textString = ""
                                                Else
                                                    varAttributes3(iattr_WKCSI).textString = SEPTEXT2(!WKNOTE, 4, 5, 48)
                                                End If
                                            End If
                                            If varAttributes3(iatHW).TagString = "XNWNOTE5" Then
                                                If IsNull(SEPTEXT2(!WKNOTE, 5, 5, 48)) = True Then
                                                    varAttributes3(iatHW).textString = ""
                                                Else
                                                    varAttributes3(iatHW).textString = SEPTEXT2(!WKNOTE, 5, 5, 48)
                                                End If
                                            End If
                                          

'**************************************************************************************
                                        Next iattr_WKCSI
                                            fillPT(1) = TblPT(1) - 16
                                            'For iFill = 3 To rnWorkImg - 1
                                                fillPT(0) = TblPT(0)
                                                fillPT(1) = fillPT(1) - 72
                                                fillPT(2) = TblPT(2)
                                                'If ((iFill Mod 2) = 0) Then
                                                    'Debug.Print "Even"
                                            For iFill = 2 To nImgFill
                                                
                                                    
                                                    Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                                        (fillPT, nImgFillDwg, 1, 1, 1, 0)
                                                    fillPT(1) = fillPT(1) - 56
                                            Next iFill
                                            
                                            If nImgFill <= 1 Then
                                                TblPT(1) = TblPT(1) - nJumpRow
                                            Else
                                                TblPT(1) = TblPT(1) - nJumpRow - ((nImgFill - 1) * 56)
                                                                                                            
                                                

                                            End If
                                        .MoveNext
                                    Next k 'rnAssWorkItem
                                    '.MoveNext
                                End With
                                'rstWorkItem.Close
                                .MoveNext
                            Next j 'rnProjAss
                            
                            End With 'rstProjAss
                   'rstProjAss.Close
                'attributeObj.Update
        
                '************************************************
                'txtRow = !drgno & Chr(9) & !oldsize
                'Set textObj = ThisDrawing.ModelSpace.AddText(txtRow, tblPT, txtHT)
                '******************************************************
                'textObj.StyleName = "cn"
                'textObj.Update
                Debug.Print "iattr_Div :" & iattr_DIV
        TblPT(1) = TblPT(1)
        .MoveNext
        
    Next i
End With
rstProjAss.Close
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                             (TblPT, nFootDwg, 1, 1, 1, 0)
''ThisDrawing.SelectionSets.Item("hdwrsched").Delete
rstProjDiv.Close
'rstProjAss.Close
'rstWorkItem.Close


'RsHDwrSet.Close
Set DB = Nothing

''Else

''ThisDrawing.SelectionSets.Item("hdwrsched").Delete
'no attribute block, inform the user
'

'MsgBox "Inserted block! Re-run this program.", vbCritical, "JWJA"

'delete the selection set
'ThisDrawing.SelectionSets.Item("TBLK").Delete

''End If

'Set MTextObj = ThisDrawing.ModelSpace.AddMText(corner, width, text)

''ThisDrawing.SelectionSets.Item("hdwrsched").Delete


Exit_writeWorkItemImg:
frmWorkItem.Hide

End Sub


