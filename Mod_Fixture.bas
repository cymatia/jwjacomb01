Attribute VB_Name = "Mod_Fixture"
Public Sub FIXTURE()

'**************************************************************
'**************************************************************
'* MOD 3/7/13 FOR PRFX NOTE
'*  MOD 9/9/14 FOR BUBBLE
'*
'**************************************************************
'**************************************************************

Dim DB As Database
Dim RstFix As Recordset
Dim strFix As String
Dim rcFix, iatFix, iatType, iatUtil As Integer
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
Dim height As Double
Dim mode As Long
Dim nHeadDwg, nRowDwg, nHeadUtilDwg, nRowUtilDwg As Variant
Dim nJumpHead, nJumpRow, nJumpHeadUtil, nJumpRowUtil As Integer
'Dim tag As String
'Dim value As String
Dim blockRefObj As AcadBlockReference
'Dim attData() As AcadObject

'Dim FilterGp(0) As Integer
'Dim FilterDt(0) As Variant
Dim BlkObj As Object
'Dim Pt1(0) As Double
'Dim Pt2(0) As Double

'Dim oSset As AcadSelectionSet



'Set DB = OpenDatabase("X:\db_misc\FIXTUREXP.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_FIX\X-FIXTURE_17_01.mdb")
If frmSched_CSI.showfood = True Then
    '******************************************************************************************************************************************************************
    '******************************************************************************************************************************************************************

   ' strFix = "SELECT FIXITEM.FXTYPE, PROFIX.PFXID, FIXITEM.FXID, PROFIX.PFXNO, PROFIX.PROJ_NO, PROFIX.PFXRM, " _
    & "PROFIX.Pfxnote, FIXITEM.FXSUPPLIER, FIXITEM.FXMODEL, FIXITEM.FXCAT, FIXITEM.FXCOLOR, FIXITEM.FXLISTPRICE, " _
    & "PROFIX.FXDISCOUNT, PROFIX.FXPRICE, PROFIX.FXTAX, FIXITEM.FXLEAD, FIXITEM.fxnote, PROFIX.FXQTY, PROFIX.FXPRNOTE, " _
    & "Contact.COMPANY, Contact_1.COMPANY, FIXITEM.FXIMGNAME, Contact.RELATION, Contact.LAST, Contact.FIRST, " _
    & "Contact.TEL_BUS, Contact.TEL_FAX, Contact.EMAIL, Contact.STREET, Contact.CITY, Contact.STATE, Contact.ZIP, " _
    & "Contact_1.RELATION, Contact_1.LAST, Contact_1.FIRST, Contact_1.TITLE, Contact_1.TEL_BUS, Contact_1.TEL_FAX, " _
    & "Contact_1.EMAIL, Contact_1.STREET, Contact_1.CITY, Contact_1.STATE, Contact_1.ZIP, PROFIX.FXEXIST, FIXITEM.FXIMGSIM, " _
    & "PROFIX.FXPRNOTE, PROFIX.COMMENT, FIXITEM.FXNAME, Trim([CONTACT].[CITY] & ', ' & [CONTACT].[STATE] & ' ' & [CONTACT].[ZIP]) AS SCSZ, " _
    & "Trim([CONTACT_1].[CITY] & ', ' & [CONTACT_1].[STATE] & ' ' & [CONTACT_1].[ZIP]) AS MCSZ FROM ((Contact RIGHT JOIN FIXITEM ON Contact.CID = FIXITEM.FXSUPPLIER) " _
    & "LEFT JOIN Contact AS Contact_1 ON FIXITEM.FXMFGR = Contact_1.CID) INNER JOIN PROFIX ON FIXITEM.FXID = PROFIX.FXID " _
    & "WHERE (((FIXITEM.FXtype) ='" & ListBox2 & "') And ((PROFIX.PROJ_NO) ='" & seleproj & "')) " _
    & "ORDER BY PROFIX.PFXNO;"
    
    '******************************************************************************************************************************************************************
    '******************************************************************************************************************************************************************
    
    '******************************************************************************************************************************************************************
    '******************************************************************************************************************************************************************
 
    'FROM (((Contact RIGHT JOIN FIXITEM ON Contact.CID = FIXITEM.FXSUPPLIER)
    'LEFT JOIN Contact AS Contact_1 ON FIXITEM.FXMFGR = Contact_1.CID)
    'LEFT JOIN UTIL ON FIXITEM.FXID = UTIL.FXID) INNER JOIN PROFIX ON
 


    'strFix = "SELECT FIXITEM.FXTYPE, PROFIX.PFXID, FIXITEM.FXID, PROFIX.PFXNO, PROFIX.PROJ_NO, PROFIX.PFXRM, PROFIX.Pfxnote, FIXITEM.FXSUPPLIER, FIXITEM.FXMODEL, FIXITEM.FXCAT, " _
     & "FIXITEM.FXCOLOR, FIXITEM.FXLISTPRICE, PROFIX.FXDISCOUNT, PROFIX.FXPRICE, PROFIX.FXTAX, FIXITEM.FXLEAD, FIXITEM.fxnote, PROFIX.FXQTY, PROFIX.FXPRNOTE, Contact.COMPANY, " _
     & "Contact_1.COMPANY, FIXITEM.FXIMGNAME, Contact.RELATION, Contact.LAST, Contact.FIRST, Contact.TEL_BUS, Contact.TEL_FAX, Contact.EMAIL, Contact.STREET, Contact.CITY, " _
     & "Contact.STATE, Contact.ZIP, Contact_1.RELATION, Contact_1.LAST, Contact_1.FIRST, Contact_1.TITLE, Contact_1.TEL_BUS, Contact_1.TEL_FAX, Contact_1.EMAIL, " _
     & "Contact_1.STREET, Contact_1.CITY, Contact_1.STATE, Contact_1.ZIP, PROFIX.FXEXIST, FIXITEM.FXIMGSIM, PROFIX.FXPRNOTE, PROFIX.COMMENT, FIXITEM.FXNAME, " _
     & "Trim([CONTACT_1].[CITY] & ', ' & [CONTACT_1].[STATE] & ' ' & [CONTACT_1].[ZIP]) AS MCSZ, Trim([CONTACT].[CITY] & ', ' & [CONTACT].[STATE] & ' ' & [CONTACT].[ZIP]) AS SCSZ, " _
     & "UTIL.BTU, UTIL.VOLT, UTIL.AMP, UTIL.WATT, UTIL.HP, UTIL.AIRTEMP, UTIL.NEMA, UTIL.LF, " _
     & "UTIL.D, UTIL.H, UTIL.NPT, UTIL.CW, UTIL.HW, UTIL.WASTE, UTIL.CONDSIZE, UTIL.PANEL , UTIL.CIRCUIT, UTIL.PH, UTIL.UTILNOTE " _
     & "FROM (((Contact RIGHT JOIN FIXITEM ON Contact.CID = FIXITEM.FXSUPPLIER) " _
     & "LEFT JOIN Contact AS Contact_1 ON FIXITEM.FXMFGR = Contact_1.CID) " _
     & "LEFT JOIN UTIL ON FIXITEM.FXID = UTIL.FXID) INNER JOIN PROFIX ON FIXITEM.FXID = PROFIX.FXID " _
     & "WHERE (((FIXITEM.FXtype) ='" & ListBox2 & "') And ((PROFIX.PROJ_NO) ='" & seleproj & "')) " _
     & "ORDER BY FIXITEM.FXTYPE, PROFIX.PFXNO;"
    '******************************************************************************************************************************************************************
    'INNER JOIN UTIL ON FIXITEM.FXID = UTIL.FXID) INNER JOIN PROFIX ON FIXITEM.FXID = PROFIX.FXID FIXITEM.FXID = PROFIX.FXID;;;;FIXITEM.FXID = PROFIX.FXID
    '******************************************************************************************************************************************************************
    '******************************************************************************************************************************************************************
     

strFix = "SELECT PROFIX.PROFIXLISTID, PROFIX.PROJ_NO, PROFIX.PFXID, PROFIX.FXID, PROFIX.PFXNO, PROFIX.PFXRM, PROFIX.FXQTY, PROFIX.FXDISCOUNT, PROFIX.FXPRICE, " _
& "PROFIX.FXTAX , PROFIX.Pfxnote, PROFIX.FXPRNOTE, FIXITEM.FXTYPE, FIXITEM.FXMODEL, FIXITEM.FXCAT, FIXITEM.FXCOLOR, FIXITEM.FXLISTPRICE, FIXITEM.FXLEAD, " _
& "FIXITEM.FXSUPPLIER, FIXITEM.FXMFGR, FIXITEM.fxnote, Project.PROJ_NAME, Contact.Relation, Contact.COMPANY, Contact.LAST, Contact.FIRST, Contact.TITLE, " _
& "Contact.TEL_BUS, Contact.TEL_FAX, Contact.EMAIL, Contact.STREET, Contact.CITY, Contact.State, Contact.ZIP, Contact.WEBADDR, FIXITEM.FXIMGNAME, " _
& "Contact_1.CID, Contact_1.Relation, Contact_1.MR, Contact_1.LAST, Contact_1.FIRST, Contact_1.COMPANY, Contact_1.TITLE, Contact_1.TEL_BUS, " _
& "Contact_1.TEL_FAX, Contact_1.EMAIL, Contact_1.STREET, Contact_1.CITY, Contact_1.State, Contact_1.ZIP, Contact_1.NOTE, Contact_1.WEBADDR, " _
& "Trim([CONTACT_1].[CITY] & ', ' & [CONTACT_1].[STATE] & ' ' & [CONTACT_1].[ZIP]) AS MCSZ, Trim([CONTACT].[CITY] & ', ' & [CONTACT].[STATE] & ' ' & [CONTACT].[ZIP]) AS SCSZ, " _
& "FIXITEM.FXNAME, PROFIX.FXEXIST, FIXITEM.FXIMGSIM, PROFIX.COMMENT, PROFIX.FXSHIPPING, PROFIX.FXUSED, FIXITEM.FXLINK, " _
& "UTIL.BTU, UTIL.VOLT, UTIL.AMP, UTIL.WATT, UTIL.HP , UTIL.AIRTEMP, UTIL.NEMA, UTIL.LF, UTIL.D, UTIL.H, UTIL.NPT, UTIL.CW, UTIL.HW, " _
& "UTIL.WASTE, UTIL.CONDSIZE, UTIL.PANEL, UTIL.CIRCUIT, UTIL.PH, UTIL.UTILNOTE, UTIL.WT " _
& "FROM (((Contact RIGHT JOIN FIXITEM ON Contact.CID = FIXITEM.FXSUPPLIER) LEFT JOIN Contact AS Contact_1 ON FIXITEM.FXMFGR = Contact_1.CID) " _
& "RIGHT JOIN (Project RIGHT JOIN PROFIX ON Project.PROJ_NO = PROFIX.PROJ_NO) ON FIXITEM.FXID = PROFIX.FXID) left JOIN UTIL ON FIXITEM.FXID = UTIL.FXID " _
& "WHERE ((FIXITEM.FXtype) = '" & frmSched_CSI.ListBox2 & "') And ((PROFIX.PROFIXLISTID) =" & frmSched_CSI.ListBox5 & ") " _
& "ORDER BY FIXITEM.FXTYPE, PROFIX.PFXNO ; "


     
    '******************************************************************************************************************************************************************
    '******************************************************************************************************************************************************************
 
    TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
    tblPT2(0) = -500: tblPT2(1) = 0: tblPT2(2) = 0
    txtHt = 4
    txtRow = ""
    mode = acAttributeModeVerify
    If frmSched_CSI.showfood = True Then
        nHeadDwg = "\\jwja-svr-10\drawfile\jvba\X_FOODEQ_HD_01.dwg"
        nRowDwg = "\\jwja-svr-10\drawfile\jvba\X_FOODEQ_row_01.dwg"
        nBubdwg = "\\jwja-svr-10\drawfile\jvba\XBUBEQ_03.dwg"
        nJumpHead = 60
        nJumpRow = 60
        InsertJBlock "\\jwja-svr-10\drawfile\jvba\X_FOODEQ_HD_01.dwg", tblPT2, "Model"
    Else
        nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XFIXHEAD01.dwg"
        nRowDwg = "\\jwja-svr-10\drawfile\jvba\XFIXROW01.dwg"
        nBubdwg = "\\jwja-svr-10\drawfile\jvba\XBUBEQ_03.dwg"
        nJumpHead = 48
        nJumpRow = 24
        '=========================================================================================
        nHeadUtilDwg = "\\jwja-svr-10\drawfile\jvba\XFIXHEAD03UTIL.dwg"
        nRowUtilDwg = "\\jwja-svr-10\drawfile\jvba\XFIXROW03UTIL.dwg"
        nJumpHeadUtil = 36
        nJumpRowUtil = 60
        '=========================================================================================
        
        InsertJBlock "\\jwja-svr-10\drawfile\jvba\XFIXHEAD01.dwg", tblPT2, "Model"
    
    End If


Else

'==============================================================================================================================
    'Trim(CONTACT_1.CITY & ', ' & CONTACT_1.STATE & ' ' & CONTACT_1.ZIP) AS SCSZ
    'strFix = "SELECT FIXITEM.FXTYPE, PROFIX.PFXID, FIXITEM.FXID, PROFIX.PFXNO, PROFIX.PROJ_NO, PROFIX.PFXRM, " _
    & "PROFIX.Pfxnote, FIXITEM.FXSUPPLIER, FIXITEM.FXMODEL, FIXITEM.FXCAT, FIXITEM.FXCOLOR, FIXITEM.FXLISTPRICE, " _
    & "PROFIX.FXDISCOUNT, PROFIX.FXPRICE, PROFIX.FXTAX, FIXITEM.FXLEAD, FIXITEM.fxnote, PROFIX.FXQTY, PROFIX.FXPRNOTE, " _
    & "Contact.COMPANY, Contact_1.COMPANY, FIXITEM.FXIMGNAME, Contact.RELATION, Contact.LAST, Contact.FIRST, " _
    & "Contact.TEL_BUS, Contact.TEL_FAX, Contact.EMAIL, Contact.STREET, Contact.CITY, Contact.STATE, Contact.ZIP, " _
    & "Contact_1.RELATION, Contact_1.LAST, Contact_1.FIRST, Contact_1.TITLE, Contact_1.TEL_BUS, Contact_1.TEL_FAX, " _
    & "Contact_1.EMAIL, Contact_1.STREET, Contact_1.CITY, Contact_1.STATE, Contact_1.ZIP, PROFIX.FXEXIST, FIXITEM.FXIMGSIM " _
    & "FROM ((Contact RIGHT JOIN FIXITEM ON Contact.CID = FIXITEM.FXSUPPLIER) " _
    & "LEFT JOIN Contact AS Contact_1 ON FIXITEM.FXMFGR = Contact_1.CID) INNER JOIN PROFIX ON FIXITEM.FXID = PROFIX.FXID " _
    & "WHERE (((FIXITEM.FXtype) ='" & ListBox2 & "') And ((PROFIX.PROJ_NO) ='" & seleproj & "')) " _
    & "ORDER BY PROFIX.PFXID, FIXITEM.FXID, PROFIX.PFXNO;"
'==============================================================================================================================
' Retired 11/3/18
'==============================================================================================================================

    'strFix = "SELECT FIXITEM.FXTYPE, PROFIX.PFXID, FIXITEM.FXID, PROFIX.PFXNO, PROFIX.PROJ_NO, PROFIX.PFXRM, " _
    & "PROFIX.Pfxnote, FIXITEM.FXSUPPLIER, FIXITEM.FXMODEL, FIXITEM.FXCAT, FIXITEM.FXCOLOR, FIXITEM.FXLISTPRICE, " _
    & "PROFIX.FXDISCOUNT, PROFIX.FXPRICE, PROFIX.FXTAX, FIXITEM.FXLEAD, FIXITEM.fxnote, PROFIX.FXQTY, PROFIX.FXPRNOTE, " _
    & "Contact.COMPANY, Contact_1.COMPANY, FIXITEM.FXIMGNAME, Contact.RELATION, Contact.LAST, Contact.FIRST, " _
    & "Contact.TEL_BUS, Contact.TEL_FAX, Contact.EMAIL, Contact.STREET, Contact.CITY, Contact.STATE, Contact.ZIP, " _
    & "Contact_1.RELATION, Contact_1.LAST, Contact_1.FIRST, Contact_1.TITLE, Contact_1.TEL_BUS, Contact_1.TEL_FAX, " _
    & "Contact_1.EMAIL, Contact_1.STREET, Contact_1.CITY, Contact_1.STATE, Contact_1.ZIP, PROFIX.FXEXIST, FIXITEM.FXIMGSIM, " _
    & "PROFIX.FXPRNOTE, PROFIX.COMMENT, FIXITEM.FXNAME, Trim([CONTACT].[CITY] & ', ' & [CONTACT].[STATE] & ' ' & [CONTACT].[ZIP]) AS SCSZ, " _
    & "Trim([CONTACT_1].[CITY] & ', ' & [CONTACT_1].[STATE] & ' ' & [CONTACT_1].[ZIP]) AS MCSZ FROM ((Contact RIGHT JOIN FIXITEM ON Contact.CID = FIXITEM.FXSUPPLIER) " _
    & "LEFT JOIN Contact AS Contact_1 ON FIXITEM.FXMFGR = Contact_1.CID) INNER JOIN PROFIX ON FIXITEM.FXID = PROFIX.FXID " _
    & "WHERE (((FIXITEM.FXtype) ='" & ListBox5 & "') And ((PROFIX.PROJ_NO) ='" & seleproj & "')) " _
    & "ORDER BY PROFIX.PFXNO;"
'==============================================================================================================================
Debug.Print "Listbox5: " & frmSched_CSI.ListBox5
Debug.Print "Listbox5: " & frmSched_CSI.ListBox2
strFix = "SELECT PROFIX.PROFIXLISTID, PROFIX.PROJ_NO, PROFIX.PFXID, PROFIX.FXID, PROFIX.PFXNO, PROFIX.PFXRM, PROFIX.FXQTY, PROFIX.FXDISCOUNT, " _
& "PROFIX.FXPRICE, PROFIX.FXTAX, PROFIX.Pfxnote, PROFIX.FXPRNOTE, FIXITEM.FXTYPE, FIXITEM.FXMODEL, FIXITEM.FXCAT, FIXITEM.FXCOLOR, FIXITEM.FXLISTPRICE, " _
& "FIXITEM.FXLEAD, FIXITEM.FXSUPPLIER, FIXITEM.FXMFGR, FIXITEM.fxnote, Project.PROJ_NAME, Contact.RELATION, Contact.COMPANY, Contact.LAST, Contact.FIRST, " _
& "Contact.TITLE, Contact.TEL_BUS, Contact.TEL_FAX, Contact.EMAIL, Contact.STREET, Contact.CITY, Contact.STATE, Contact.ZIP, Contact.WEBADDR, " _
& "FIXITEM.FXIMGNAME, Contact_1.CID, Contact_1.RELATION, Contact_1.MR, Contact_1.LAST, Contact_1.FIRST, Contact_1.COMPANY, Contact_1.TITLE, " _
& "Contact_1.TEL_BUS, Contact_1.TEL_FAX, Contact_1.EMAIL, Contact_1.STREET, Contact_1.CITY, Contact_1.STATE, Contact_1.ZIP, Contact_1.NOTE, Contact_1.WEBADDR, " _
& "Trim([CONTACT_1].[CITY] & ', ' & [CONTACT_1].[STATE] & ' ' & [CONTACT_1].[ZIP]) AS MCSZ, Trim([CONTACT].[CITY] & ', ' & [CONTACT].[STATE] & ' ' & [CONTACT].[ZIP]) AS SCSZ, " _
& "FIXITEM.FXNAME, PROFIX.FXEXIST, FIXITEM.FXIMGSIM, PROFIX.COMMENT, PROFIX.FXSHIPPING, PROFIX.FXUSED, FIXITEM.FXLINK " _
& "FROM ((Contact RIGHT JOIN FIXITEM ON Contact.CID = FIXITEM.FXSUPPLIER) LEFT JOIN Contact AS Contact_1 ON FIXITEM.FXMFGR = Contact_1.CID) " _
& "RIGHT JOIN (Project RIGHT JOIN PROFIX ON Project.PROJ_NO = PROFIX.PROJ_NO) ON FIXITEM.FXID = PROFIX.FXID " _
& "WHERE ((PROFIX.PROFIXLISTID) =" & frmSched_CSI.ListBox5 & ")  and ((FIXITEM.FXtype) = '" & frmSched_CSI.ListBox2 & "')" _
& "ORDER BY FIXITEM.FXTYPE, PROFIX.PFXNO;"
'==============================================================================================================================
Debug.Print "Listbox5: " & frmSched_CSI.ListBox5
    TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
    tblPT2(0) = -500: tblPT2(1) = 0: tblPT2(2) = 0
    txtHt = 4
    txtRow = ""
    mode = acAttributeModeVerify
    If frmSched_CSI.ShowImage = True Then
        nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XFIXHEAD03.dwg"    'MOD 9/9/14
        nRowDwg = "\\jwja-svr-10\drawfile\jvba\XFIXROW03.dwg"  'MOD 9/9/14
        nBubdwg = "\\jwja-svr-10\drawfile\jvba\XBUBEQ_03.dwg"  'MOD 9/9/14
        nJumpHead = 60
        nJumpRow = 60
        nHeadUtilDwg = "\\jwja-svr-10\drawfile\jvba\XFIXHEAD03UTIL.dwg"
        nRowUtilDwg = "\\jwja-svr-10\drawfile\jvba\XFIXROW03UTIL.dwg"
        nJumpHeadUtil = 36 '3x12=36
        nJumpRowUtil = 60
        InsertJBlock "\\jwja-svr-10\drawfile\jvba\XFIXHEAD03.dwg", tblPT2, "Model"
    Else
        nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XFIXHEAD01.dwg"
        nRowDwg = "\\jwja-svr-10\drawfile\jvba\XFIXROW01.dwg"
        nBubdwg = "\\jwja-svr-10\drawfile\jvba\XBUBEQ_03.dwg"  'MOD 9/9/14
        nJumpHead = 48
        nJumpRow = 24
        nHeadUtilDwg = "\\jwja-svr-10\drawfile\jvba\XFIXHEAD03UTIL.dwg"
        nRowUtilDwg = "\\jwja-svr-10\drawfile\jvba\XFIXROW03UTIL.dwg"
        nJumpHeadUtil = 36
        nJumpRowUtil = 60
        InsertJBlock "\\jwja-svr-10\drawfile\jvba\XFIXHEAD01.dwg", tblPT2, "Model"
    
    End If

End If
'ThisDrawing.ModelSpace.InsertBlock tblPT, "H:\drawfile\jvba\XFIXHEAD01.dwg", 1, 1, 1, 0
'InsertJBlock "H:\drawfile\jvba\XDWGTRADE01.dwg", tblPT2, "Model"
'InsertJBlock nHeadDwg, tblPT2, "Model"
'InsertJBlock "e:\drawfile\jvba\XHDWRFOOT02.dwg", tblPT2, "Model"
'tblPT(1) = tblPT(1) - 48

Set RstFix = DB.OpenRecordset(strFix)
rcFix = RstFix.RecordCount
Debug.Print rcFix
RstFix.MoveLast
RstFix.MoveFirst
rcFix = RstFix.RecordCount
Debug.Print rcFix
With RstFix
Set blockRefObj2 = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, nHeadDwg, 1, 1, 1, 0)
Dim varAttributes2 As Variant
varAttributes2 = blockRefObj2.GetAttributes
For iatType = LBound(varAttributes2) To UBound(varAttributes2)

Debug.Print varAttributes2(iatType).TagString
    If varAttributes2(iatType).TagString = "XTYPE" Then
        If IsNull(!FXtype) = True Then
            varAttributes2(iatType).textString = ""
        Else
            varAttributes2(iatType).textString = !FXtype & " SCHEDULE"
            Debug.Print varAttributes2(iatType).textString
        End If
    End If
    If varAttributes2(iatType).TagString = "XPROJ_NO" Then
        If IsNull(!PROJ_NO) = True Then
            varAttributes2(iatType).textString = ""
        Else
            varAttributes2(iatType).textString = !PROJ_NO
            Debug.Print varAttributes2(iatType).textString
        End If
    End If
    If varAttributes2(iatType).TagString = "XDATE" Then
        'If IsNull(!fxtype) = True Then
            'varAttributes2(iatType).TextString = ""
        'Else
            varAttributes2(iatType).textString = Date
            Debug.Print varAttributes2(iatType).textString
        'End If
    End If

Next iatType
    TblPT(1) = TblPT(1) - nJumpHead
    For i = 1 To rcFix
        ' xHDWRSETID = !HDWRSETID
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, nRowDwg, 1, 1, 1, 0)
        Dim varAttributes As Variant
        varAttributes = blockRefObj.GetAttributes
    '***************************************************************
    If !FXIMGNAME <> "" Then
    If frmSched_CSI.ShowImage = True Then

    Dim imgPT(0 To 2) As Double
    Dim scalefactor As Double
    Dim rotAngleInDegree As Double, rotAngle As Double
    'Dim imageName As String
    Dim imageName As String

    Dim raster As AcadRasterImage
    Dim imgheight As Variant
    Dim imgwidth As Variant

    'imageName = "C:\AutoCAD\sample\downtown.jpg"
    imageName = !FXIMGNAME
    Debug.Print imageName
    'insertionPoint(0) = 2#: insertionPoint(1) = 2#: insertionPoint(2) = 0#
    imgPT(0) = TblPT(0) + 36: imgPT(1) = TblPT(1) - 54: imgPT(2) = TblPT(2)
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
    End If
    '***************************************************************
        For iatFix = LBound(varAttributes) To UBound(varAttributes)
        
        Debug.Print varAttributes(iatFix).TagString
            If varAttributes(iatFix).TagString = "XKEY" Then 'XKEY
                If IsNull(![PFXNO]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = ![PFXNO]
                    Debug.Print varAttributes(iatFix).textString
    '***************************************************************
    '*      INSERT BUBBLE
    '***************************************************************
                    TblPT(0) = TblPT(0) - 48
                    Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                 (TblPT, nBubdwg, 1, 1, 1, 0)
                    Dim varAttributesB As Variant
                    varAttributesB = blockRefObj.GetAttributes
                    varAttributesB(0).textString = ![PFXNO]
                    
                    TblPT(0) = TblPT(0) + 48
    
    '***************************************************************
                End If
            End If
            If varAttributes(iatFix).TagString = "XMODEL1" Then 'XMODEL1
                If IsNull(!FXMODEL) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXMODEL, 1, 2, 16)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXMODEL, 1, 2, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XMODEL2" Then 'XMODEL2
                If IsNull(!FXMODEL) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXMODEL, 2, 2, 16)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXMODEL, 2, 2, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XNAME" Then 'XNAME
                If IsNull(!FXNAME) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !FXNAME
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XCAT" Then 'XCAT
                If IsNull(!FXCAT) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !FXCAT
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XFIN1" Then 'XFIN1
                If IsNull(!FXCOLOR) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXCOLOR, 1, 2, 7)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXCOLOR, 1, 2, 7)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XFIN2" Then 'XFIN2
                If IsNull(!FXCOLOR) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXCOLOR, 2, 2, 7)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXCOLOR, 2, 2, 7)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            
            If varAttributes(iatFix).TagString = "XMCOMP" Then 'XMCOMP
                If IsNull(![CONTACT_1.COMPANY]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = Mid(![CONTACT_1.COMPANY], 1, 26)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XMSTREET" Then 'XMSTREET
                If IsNull(![CONTACT_1.STREET]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = ![CONTACT_1.STREET]
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XMCSZ" Then 'XMCSZ
                If !MCSZ = "," Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !MCSZ
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XMTEL" Then 'XMTEL
                If IsNull(![CONTACT_1.TEL_BUS]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = Format(![CONTACT_1.TEL_BUS], "###-###-####")
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XSCOMP" Then 'XSCOMP
                If IsNull(![CONTACT.COMPANY]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = Mid(![CONTACT.COMPANY], 1, 26)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XSSTREET" Then 'XSSTREET
                If IsNull(![CONTACT.STREET]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = ![CONTACT.STREET]
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XSCSZ" Then 'XSCSZ
                If !SCSZ = "," Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !SCSZ
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XSTEL" Then 'XSTEL
                If IsNull(![CONTACT.TEL_BUS]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = Format(![CONTACT.TEL_BUS], "###-###-####")
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XNOTE1" Then 'XNOTE1 FXPRNOTE
                If IsNull(!FXPRNOTE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXPRNOTE, 1, 3, 16)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXPRNOTE, 1, 3, 30)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XNOTE2" Then 'XNOTE2
                If IsNull(!FXPRNOTE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXPRNOTE, 2, 3, 16)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXPRNOTE, 2, 3, 30)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XNOTE3" Then 'XNOTE3
                If IsNull(!FXPRNOTE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXPRNOTE, 3, 3, 16)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXPRNOTE, 3, 3, 30)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
   '************************************************************************************************
   '************************************************************************************************
            If varAttributes(iatFix).TagString = "XFXNOTE1" Then 'XNOTE3
                If IsNull(!fxnote) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    If showfood = True Then
                    varAttributes(iatFix).textString = SEPTEXT2(!fxnote, 1, 3, 120)
                    Else
                    varAttributes(iatFix).textString = SEPTEXT2(!fxnote, 1, 3, 30)
                    End If
                    
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            
            If varAttributes(iatFix).TagString = "XFXNOTE2" Then 'XNOTE3
                If IsNull(!fxnote) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    If showfood = True Then
                    varAttributes(iatFix).textString = SEPTEXT2(!fxnote, 2, 3, 120)
                    Else
                    varAttributes(iatFix).textString = SEPTEXT2(!fxnote, 2, 3, 30)
                    End If
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            
            If varAttributes(iatFix).TagString = "XFXNOTE3" Then 'XNOTE3
                If IsNull(!fxnote) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    If showfood = True Then
                    varAttributes(iatFix).textString = SEPTEXT2(!fxnote, 3, 3, 120)
                    Else
                    varAttributes(iatFix).textString = SEPTEXT2(!fxnote, 3, 3, 30)
                    End If
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If

Food:
   '************************************************************************************************
   '************************************************************************************************
   '************************************ food equipments *******************************************
   '************************************************************************************************
   '************************************************************************************************
        If frmSched_CSI.showfood = True Then
            If varAttributes(iatFix).TagString = "BTU" Then 'BTU
                If IsNull(!BTU) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = Format(!BTU, "##,#")
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "NPT" Then 'BTU
                If IsNull(!NPT) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !NPT
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "CW" Then 'BTU
                If IsNull(!CW) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !CW
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            
            If varAttributes(iatFix).TagString = "HW" Then 'HW
                If IsNull(!HW) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !HW
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "WASTE" Then 'WASTE
                If IsNull(!WASTE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !WASTE
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "HP" Then 'HP
                If IsNull(!HP) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !HP
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "WATT" Then 'WATT
                If IsNull(!WATT) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = Format(!WATT, "##,#")
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "AMP" Then 'AMP
                If IsNull(!AMP) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !AMP
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "VOLT" Then 'VOLT
                If IsNull(!VOLT) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !VOLT
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "PH" Then 'PH
                If IsNull(!PH) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !PH
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "NEMA" Then 'NEMA
                If IsNull(!NEMA) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !NEMA
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "PANEL" Then 'PANEL
                If IsNull(!PANEL) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !PANEL
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "CIRCUIT" Then 'CIRCUIT
                If IsNull(!CIRCUIT) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !CIRCUIT
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "AIRTEMP" Then 'AIRTEMP
                If IsNull(!AIRTEMP) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !AIRTEMP
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "LF" Then 'LF
                If IsNull(!LF) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !LF
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "D" Then 'D
                If IsNull(!D) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !D
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "H" Then 'H
                If IsNull(!H) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !H
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "UTILNOTE1" Then 'UTILNOTE1
                If IsNull(!UTILNOTE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = SEPTEXT2(!UTILNOTE, 1, 3, 24)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "UTILNOTE2" Then 'UTILNOTE2
                If IsNull(!UTILNOTE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = SEPTEXT2(!UTILNOTE, 2, 3, 24)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "UTILNOTE3" Then 'UTILNOTE3
                If IsNull(!UTILNOTE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = SEPTEXT2(!UTILNOTE, 3, 3, 24)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            
       End If ' IF SHOWFOOD = TRUE THEN
            
            
            
            
            
            
   '************************************************************************************************
   '************************************************************************************************
   '************************************END OF food equipments *************************************
   '************************************************************************************************
   '************************************************************************************************
            
            
            
            
            'Dim selectionSet1 As AcadSelectionSet
            'Set selectionSet1 = ThisDrawing.SelectionSets. _
                        Add("NewSelectionSet")

            'ssImg = ThisDrawing.SelectionSets.Item("ImgSet").Add
            'Set rasterObj.Name = ""
            'rasterObj = ThisDrawing.modelspace.ObjectName "cadimg"
            'Set rasterObj.ImageFile = !FXIMGNAME
            
        Next iatFix
        '==================================================================================================
        '       Utility
        '==================================================================================================
        'nHeadUtilDwg = "\\jwja-svr-10\drawfile\jvba\XFIXHEAD03UTIL.dwg"
        'nRowUtilDwg = "\\jwja-svr-10\drawfile\jvba\XFIXROW03UTIL.dwg"
        'nJumpHeadUtil = 36
        'nJumpRowUtil = 60

        Dim rstUtil As Recordset
        strUtil = "SELECT UTIL.UTILID, UTIL.FXID, UTIL.BTU, UTIL.VOLT, UTIL.AMP, UTIL.WATT, UTIL.HP, UTIL.AIRTEMP, UTIL.NEMA, UTIL.LF, " _
        & "UTIL.D, UTIL.H, UTIL.NPT, UTIL.CW, UTIL.HW, UTIL.WASTE, UTIL.CONDSIZE, UTIL.PANEL, UTIL.CIRCUIT, UTIL.PH, UTIL.UTILNOTE, UTIL.WT " _
        & "FROM UTIL " _
        & "WHERE (UTIL.FXID)=" & RstFix!fxid

        Set rstUtil = DB.OpenRecordset(strUtil)
        hasutil = 0
 
         With rstUtil
            If rstUtil.EOF Then
                GoTo NoUtil
                hasutil = 0
            Else
                TblPT(1) = TblPT(1) - nJumpRow
             
                ' xHDWRSETID = !HDWRSETID
                Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                             (TblPT, nHeadUtilDwg, 1, 1, 1, 0)
                TblPT(1) = TblPT(1) - nJumpHeadUtil
                
                Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                             (TblPT, nRowUtilDwg, 1, 1, 1, 0)
                Dim varAttributesUtil As Variant
                varAttributesUtil = blockRefObj.GetAttributes
                For iatUtil = LBound(varAttributesUtil) To UBound(varAttributesUtil)
                '***************************************************************************************
                    If varAttributesUtil(iatUtil).TagString = "BTU" Then 'BTU
                        If IsNull(!BTU) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = Format(!BTU, "##,#")
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "NPT" Then 'BTU
                        If IsNull(!NPT) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !NPT
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "CW" Then 'BTU
                        If IsNull(!CW) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !CW
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    
                    If varAttributesUtil(iatUtil).TagString = "HW" Then 'HW
                        If IsNull(!HW) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !HW
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "WASTE" Then 'WASTE
                        If IsNull(!WASTE) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !WASTE
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "HP" Then 'HP
                        If IsNull(!HP) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !HP
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "WATT" Then 'WATT
                        If IsNull(!WATT) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = Format(!WATT, "##,#")
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "AMP" Then 'AMP
                        If IsNull(!AMP) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !AMP
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "VOLT" Then 'VOLT
                        If IsNull(!VOLT) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !VOLT
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "PH" Then 'PH
                        If IsNull(!PH) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !PH
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "NEMA" Then 'NEMA
                        If IsNull(!NEMA) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !NEMA
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "PANEL" Then 'PANEL
                        If IsNull(!PANEL) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !PANEL
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "CIRCUIT" Then 'CIRCUIT
                        If IsNull(!CIRCUIT) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !CIRCUIT
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "AIRTEMP" Then 'AIRTEMP
                        If IsNull(!AIRTEMP) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !AIRTEMP
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "LF" Then 'LF
                        If IsNull(!LF) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !LF
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "D" Then 'D
                        If IsNull(!D) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !D
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "H" Then 'H
                        If IsNull(!H) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !H
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "WT" Then 'WT
                        If IsNull(!WT) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = !WT
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "UTILNOTE1" Then 'UTILNOTE1
                        If IsNull(!UTILNOTE) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = SEPTEXT2(!UTILNOTE, 1, 3, 56)
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "UTILNOTE2" Then 'UTILNOTE2
                        If IsNull(!UTILNOTE) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = SEPTEXT2(!UTILNOTE, 2, 3, 56)
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
                    If varAttributesUtil(iatUtil).TagString = "UTILNOTE3" Then 'UTILNOTE3
                        If IsNull(!UTILNOTE) = True Then
                            varAttributesUtil(iatUtil).textString = ""
                        Else
                            varAttributesUtil(iatUtil).textString = SEPTEXT2(!UTILNOTE, 3, 3, 56)
                            Debug.Print varAttributesUtil(iatUtil).textString
                        End If
                    End If
               
                '***************************************************************************************
                Next iatUtil
            End If
        End With
        TblPT(1) = TblPT(1) - nJumpRowUtil
        hasutil = 1
        rstUtil.Close
 
 
'==================================================================================================
'==================================================================================================
NoUtil:
        
        
        .MoveNext
        If hasutil = 1 Then
            TblPT(1) = TblPT(1)
        Else
            TblPT(1) = TblPT(1) - nJumpRow
        End If
        
    Next i
End With

frmSched_CSI.Hide

End Sub

Public Sub FIXTUREFOOD()

'**************************************************************
'**************************************************************
'* MOD 3/7/13 FOR PRFX NOTE
'*
'*
'**************************************************************
'**************************************************************

Dim DB As Database
Dim RstFix As Recordset
Dim strFix As String
Dim rcFix, iatFix, iatType As Integer
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
Dim height As Double
Dim mode As Long
Dim nHeadDwg, nRowDwg As Variant
Dim nJumpHead, nJumpRow As Integer
'Dim tag As String
'Dim value As String
Dim blockRefObj As AcadBlockReference
'Dim attData() As AcadObject

'Dim FilterGp(0) As Integer
'Dim FilterDt(0) As Variant
Dim BlkObj As Object
'Dim Pt1(0) As Double
'Dim Pt2(0) As Double

'Dim oSset As AcadSelectionSet



Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_FIX\X-FIXTURE_17_01.mdb")
'*************************************************************************************************************
'strFix = "SELECT FIXITEM.FXTYPE, PROFIX.PFXID, FIXITEM.FXID, PROFIX.PFXNO, PROFIX.PROJ_NO, PROFIX.PFXRM, " _
& "PROFIX.Pfxnote, FIXITEM.FXSUPPLIER, FIXITEM.FXMODEL, FIXITEM.FXCAT, FIXITEM.FXCOLOR, FIXITEM.FXLISTPRICE, " _
& "PROFIX.FXDISCOUNT, PROFIX.FXPRICE, PROFIX.FXTAX, FIXITEM.FXLEAD, FIXITEM.fxnote, PROFIX.FXQTY, PROFIX.FXPRNOTE, " _
& "Contact.COMPANY, Contact_1.COMPANY, FIXITEM.FXIMGNAME, Contact.RELATION, Contact.LAST, Contact.FIRST, " _
& "Contact.TEL_BUS, Contact.TEL_FAX, Contact.EMAIL, Contact.STREET, Contact.CITY, Contact.STATE, Contact.ZIP, " _
& "Contact_1.RELATION, Contact_1.LAST, Contact_1.FIRST, Contact_1.TITLE, Contact_1.TEL_BUS, Contact_1.TEL_FAX, " _
& "Contact_1.EMAIL, Contact_1.STREET, Contact_1.CITY, Contact_1.STATE, Contact_1.ZIP, PROFIX.FXEXIST, FIXITEM.FXIMGSIM, " _
& "PROFIX.FXPRNOTE, PROFIX.COMMENT, FIXITEM.FXNAME, Trim([CONTACT].[CITY] & ', ' & [CONTACT].[STATE] & ' ' & [CONTACT].[ZIP]) AS SCSZ, " _
& "Trim([CONTACT_1].[CITY] & ', ' & [CONTACT_1].[STATE] & ' ' & [CONTACT_1].[ZIP]) AS MCSZ FROM ((Contact RIGHT JOIN FIXITEM ON Contact.CID = FIXITEM.FXSUPPLIER) " _
& "LEFT JOIN Contact AS Contact_1 ON FIXITEM.FXMFGR = Contact_1.CID) INNER JOIN PROFIX ON FIXITEM.FXID = PROFIX.FXID " _
& "WHERE (((FIXITEM.FXtype) ='" & ListBox2 & "') And ((PROFIX.PROJ_NO) ='" & seleproj & "')) " _
& "ORDER BY PROFIX.PFXNO;"
' THIS OLD
'*************************************************************************************************************









TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -500: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify
If frmSched_CSI.ShowImage = True Then
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XFIXHEAD02.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\XFIXROW02.dwg"
    nJumpHead = 60
    nJumpRow = 60
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\XFIXHEAD02.dwg", tblPT2, "Model"
Else
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XFIXHEAD01.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\XFIXROW01.dwg"
    nJumpHead = 48
    nJumpRow = 24
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\XFIXHEAD01.dwg", tblPT2, "Model"

End If
'ThisDrawing.ModelSpace.InsertBlock tblPT, "H:\drawfile\jvba\XFIXHEAD01.dwg", 1, 1, 1, 0
'InsertJBlock "H:\drawfile\jvba\XDWGTRADE01.dwg", tblPT2, "Model"
'InsertJBlock nHeadDwg, tblPT2, "Model"
'InsertJBlock "e:\drawfile\jvba\XHDWRFOOT02.dwg", tblPT2, "Model"
'tblPT(1) = tblPT(1) - 48

Set RstFix = DB.OpenRecordset(strFix)
RstFix.MoveLast
RstFix.MoveFirst
rcFix = RstFix.RecordCount
Debug.Print rcFix
With RstFix
Set blockRefObj2 = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, nHeadDwg, 1, 1, 1, 0)
Dim varAttributes2 As Variant
varAttributes2 = blockRefObj2.GetAttributes
For iatType = LBound(varAttributes2) To UBound(varAttributes2)

Debug.Print varAttributes2(iatType).TagString
    If varAttributes2(iatType).TagString = "XTYPE" Then
        If IsNull(!FXtype) = True Then
            varAttributes2(iatType).textString = ""
        Else
            varAttributes2(iatType).textString = !FXtype & " SCHEDULE"
            Debug.Print varAttributes2(iatType).textString
        End If
    End If
    If varAttributes2(iatType).TagString = "XPROJ_NO" Then
        If IsNull(!PROJ_NO) = True Then
            varAttributes2(iatType).textString = ""
        Else
            varAttributes2(iatType).textString = !PROJ_NO
            Debug.Print varAttributes2(iatType).textString
        End If
    End If
    If varAttributes2(iatType).TagString = "XDATE" Then
        'If IsNull(!fxtype) = True Then
            'varAttributes2(iatType).TextString = ""
        'Else
            varAttributes2(iatType).textString = Date
            Debug.Print varAttributes2(iatType).textString
        'End If
    End If

Next iatType
    TblPT(1) = TblPT(1) - nJumpHead
    For i = 1 To rcFix
        ' xHDWRSETID = !HDWRSETID
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, nRowDwg, 1, 1, 1, 0)
        Dim varAttributes As Variant
        varAttributes = blockRefObj.GetAttributes
    '***************************************************************
    If !FXIMGNAME <> "" Then
    If frmSched_CSI.ShowImage = True Then

    Dim imgPT(0 To 2) As Double
    Dim scalefactor As Double
    Dim rotAngleInDegree As Double, rotAngle As Double
    'Dim imageName As String
    Dim imageName As String

    Dim raster As AcadRasterImage
    Dim imgheight As Variant
    Dim imgwidth As Variant

    'imageName = "C:\AutoCAD\sample\downtown.jpg"
    imageName = !FXIMGNAME
    Debug.Print imageName
    'insertionPoint(0) = 2#: insertionPoint(1) = 2#: insertionPoint(2) = 0#
    imgPT(0) = TblPT(0) + 36: imgPT(1) = TblPT(1) - 54: imgPT(2) = TblPT(2)
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
    End If
    '***************************************************************
        For iatFix = LBound(varAttributes) To UBound(varAttributes)
        
        Debug.Print varAttributes(iatFix).TagString
            If varAttributes(iatFix).TagString = "XKEY" Then 'XKEY
                If IsNull(![PFXNO]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = ![PFXNO]
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XMODEL1" Then 'XMODEL1
                If IsNull(!FXMODEL) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXMODEL, 1, 2, 16)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXMODEL, 1, 2, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XMODEL2" Then 'XMODEL2
                If IsNull(!FXMODEL) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXMODEL, 2, 2, 16)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXMODEL, 2, 2, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XNAME" Then 'XNAME
                If IsNull(!FXNAME) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !FXNAME
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XCAT" Then 'XCAT
                If IsNull(!FXCAT) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !FXCAT
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XFIN1" Then 'XFIN1
                If IsNull(!FXCOLOR) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXCOLOR, 1, 2, 7)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXCOLOR, 1, 2, 7)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XFIN2" Then 'XFIN2
                If IsNull(!FXCOLOR) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXCOLOR, 2, 2, 7)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXCOLOR, 2, 2, 7)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            
            If varAttributes(iatFix).TagString = "XMCOMP" Then 'XMCOMP
                If IsNull(![CONTACT_1.COMPANY]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = Mid(![CONTACT_1.COMPANY], 1, 26)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XMSTREET" Then 'XMSTREET
                If IsNull(![CONTACT_1.STREET]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = ![CONTACT_1.STREET]
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XMCSZ" Then 'XMCSZ
                If !MCSZ = "," Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !MCSZ
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XMTEL" Then 'XMTEL
                If IsNull(![CONTACT_1.TEL_BUS]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = Format(![CONTACT_1.TEL_BUS], "###-###-####")
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XSCOMP" Then 'XSCOMP
                If IsNull(![CONTACT.COMPANY]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = Mid(![CONTACT.COMPANY], 1, 26)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XSSTREET" Then 'XSSTREET
                If IsNull(![CONTACT.STREET]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = ![CONTACT.STREET]
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XSCSZ" Then 'XSCSZ
                If !SCSZ = "," Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !SCSZ
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XSTEL" Then 'XSTEL
                If IsNull(![CONTACT.TEL_BUS]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = Format(![CONTACT.TEL_BUS], "###-###-####")
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XNOTE1" Then 'XNOTE1 FXPRNOTE
                If IsNull(!FXPRNOTE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXPRNOTE, 1, 3, 16)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXPRNOTE, 1, 3, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XNOTE2" Then 'XNOTE2
                If IsNull(!FXPRNOTE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXPRNOTE, 2, 3, 16)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXPRNOTE, 2, 3, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XNOTE3" Then 'XNOTE3
                If IsNull(!FXPRNOTE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    'varAttributes(iatFix).TextString = SEPTEXT(!FXPRNOTE, 3, 3, 16)
                    varAttributes(iatFix).textString = SEPTEXT2(!FXPRNOTE, 3, 3, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
   '************************************************************************************************
   '************************************************************************************************
            If varAttributes(iatFix).TagString = "XFXNOTE1" Then 'XNOTE3
                If IsNull(!fxnote) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = SEPTEXT2(!fxnote, 1, 3, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            
            If varAttributes(iatFix).TagString = "XFXNOTE2" Then 'XNOTE3
                If IsNull(!fxnote) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = SEPTEXT2(!fxnote, 2, 3, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            
            If varAttributes(iatFix).TagString = "XFXNOTE3" Then 'XNOTE3
                If IsNull(!fxnote) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = SEPTEXT2(!fxnote, 3, 3, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
   
   '************************************************************************************************
   '************************************************************************************************
            
            'Dim selectionSet1 As AcadSelectionSet
            'Set selectionSet1 = ThisDrawing.SelectionSets. _
                        Add("NewSelectionSet")

            'ssImg = ThisDrawing.SelectionSets.Item("ImgSet").Add
            'Set rasterObj.Name = ""
            'rasterObj = ThisDrawing.modelspace.ObjectName "cadimg"
            'Set rasterObj.ImageFile = !FXIMGNAME
            
        Next iatFix
        TblPT(1) = TblPT(1) - nJumpRow
        .MoveNext
    Next i
End With

UserForm2.Hide

End Sub


