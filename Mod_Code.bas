Attribute VB_Name = "Mod_Code"
Public Sub ShowCode()
frmCode.show
End Sub
Public Sub writeCode(mShow As Variant)
'############################################################################
'############################################################################
'#          11/15/12 Working JIN
'#
'#
'#
'#
'#
'############################################################################
'############################################################################

Dim DB As Database
Dim RstProjCode, rstProjMuni, rstRefCode, rstBisRow, rstOCC, rstConType, rstExtElem0, rstExtElem1 As Recordset
Dim strProjCode, strProjMuni, strRefCode, strBisRow, strOCC, strConType, strExtElem0, strExtElem1 As String
Dim nRefCode, nBISrow, nOCCrow As Integer    'ProjSet no.
Dim iattr_CODE, iattr_refcode, iattr_BISrow, iattr_OCC, iattr_ConType, iattr_ExtElem0, iattr_ExtElem1 As Integer
'Dim rcArray As Variant
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
'Dim tmpArrayAss(0 To 1) As Double

'Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
'Dim height As Double
Dim mode As Long
'Dim prompt As String

Dim nCODEHeadDwg, nREFCODEDwg, nDOBHDDwg, nDOBRowDwg, nCODEBCHDDwg As Variant
Dim nJumpHead, nJumpREFCODEROW, nJumpDOBHD, nJumpBISRow As Integer
Dim blockRefObj As AcadBlockReference
Dim BlkObj As Object
Dim k, l As Integer

'frmWorkItem.ListBox1.Clear
CreateCNstyle
TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify
    If mShow = False Then
        nCODEHeadDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_HD01.dwg"
    Else
        nCODEHeadDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_HD02.dwg" '8/25/14
    End If
    nREFCODEDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_REFCODE_ROW01.dwg"
    nDOBHDDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_DOB_HD01.dwg"
    nDOBRowDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_BIS_ROW01.dwg"
    nCODEBCHDDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_BC_HD01.dwg"
    nOCCTYPEROWDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_OCC_ROW01.dwg"
    nCONTYPEHDDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_CONTYPE_HD01.dwg"
    nCONTYPEROWDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_CONTYPE_ROW01.dwg"
    nAHHDDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_AH_HD01.dwg"
    nAHROWDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_AH_ROW01.dwg"
    
    nELEMHDDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_CON_ELEM_HD01.dwg"
    nELEM0Dwg = "\\jwja-svr-10\drawfile\jvba\XCODE_CON_ELEM0_ROW01.dwg"
    nELEM1Dwg = "\\jwja-svr-10\drawfile\jvba\XCODE_CON_ELEM1_ROW01.dwg"
    If mShow = False Then
    nJumpHead = 204 '17 X 12
    Else
    nJumpHead = 240 '20 X 12  '8/25/14
    End If
    nJumpREFCODEROW = 12 '1 X 12
    nJumpDOBHD = 18
    nJumpBISRow = 24 '2 X 12
    nJumpOCCTYPEHD = 42 '2 X 12
    nJumpOCCTYPEROW = 12 '2 X 12
    nJumpCONTYPEHD = 24
    nJumpCONTYPEROW = 12
    nJumpAHHD = 36
    nJumpAHROW = 12
    
    nJumpCONELEMHD = 36
    nJumpCONELEM0ROW = 12
    nJumpCONELEM1ROW = 12
    
    
    'InsertJBlock "H:\drawfile\jvba\XCODE_HD01.dwg", tblPT, "Model"

'tblPT(1) = tblPT(1) - nJumpHead

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_OUTSPEC\x-cl_org_outspec_02.mdb")
xProj_no = frmCode.seleproj
Debug.Print xAssID
strProjCode = "SELECT PROJCODE.PROJCODEID, PROJCODE.PROJ_NO, PROJCODE.OCCLOADID, " _
& "Project.PROJ_NAME, Project.PROJ_STNO, Project.PROJ_ADDR, Project.PROJ_CITY, " _
& "Project.PROJ_STATE, Project.PROJ_ZIP, Project.projdesc, Project.PROJ_TYPE, Project.TYPE, " _
& "Project.STATUS, Project.CID, PROJCODE.BLOCK, PROJCODE.LOT, PROJCODE.SECTION, " _
& "PROJCODE.MUNI, PROJCODE.ZONE, PROJCODE.LOTSIZE, PROJCODE.COVERAGE_EXISTING, " _
& "PROJCODE.COVERAGE_NEW, PROJCODE.COMMBD, PROJCODE.BIN, PROJCODE.STORY, PROJCODE.STORY_NW, PROJCODE.OVZONE, " _
& "PROJCODE.ZONEMAP, PROJCODE.BLDGHT_EX, PROJCODE.BLDGHT_NW, PROJCODE.AREA_EX, PROJCODE.AREA_ADD, PROJCODE.AREA_NW, PROJCODE.AREA_TOTAL, " _
& "PROJCODE.DU_EX, PROJCODE.DU_ADD, PROJCODE.DU_NW, PROJCODE.DU_TOTAL, " _
& "Contact.COMPANY, Contact.STREET, Contact.CITY, Contact.STATE, Contact.ZIP, Contact.WEBADDR, Trim([city] & ', ' & [state] & ' ' & [zip]) AS csz " _
& "FROM (PROJCODE INNER JOIN Project ON PROJCODE.PROJ_NO = Project.PROJ_NO) LEFT JOIN Contact ON PROJCODE.MUNI = Contact.CID " _
& "WHERE (((PROJCODE.PROJ_NO)='" & xProj_no & "'));"

'strProjCode = "SELECT PROJCODE.PROJCODEID, PROJCODE.PROJ_NO, PROJCODE.OCCLOADID, " _
& "Project.PROJ_NAME, Project.PROJ_STNO, Project.PROJ_ADDR, Project.PROJ_CITY, " _
& "Project.PROJ_STATE, Project.PROJ_ZIP, Project.projdesc, Project.PROJ_TYPE, " _
& "Project.TYPE, Project.STATUS, Project.CID, PROJCODE.BLOCK, PROJCODE.LOT, " _
& "PROJCODE.SECTION, PROJCODE.MUNI, PROJCODE.ZONE, PROJCODE.LOTSIZE, " _
& "PROJCODE.COVERAGE_EXISTING, PROJCODE.COVERAGE_NEW, PROJCODE.COMMBD, PROJCODE.BIN, " _
& "PROJCODE.STORY, PROJCODE.OVZONE, PROJCODE.ZONEMAP, PROJCODE.BLDGHT_EX, PROJCODE.BLDGHT_NW, " _
& "Contact.COMPANY, Contact.STREET, Contact.CITY, Contact.STATE, Contact.ZIP, Contact.WEBADDR, Trim([city] & ', ' & [state] & ' ' & [zip]) AS csz " _
& "FROM (PROJCODE INNER JOIN Project ON PROJCODE.PROJ_NO = Project.PROJ_NO) INNER JOIN Contact ON PROJCODE.MUNI = Contact.CID " _
& "WHERE (((PROJCODE.PROJ_NO)='" & xProj_no & "'));"


Set RstProjCode = DB.OpenRecordset(strProjCode)
Debug.Print "REC Count: " & RstProjCode.RecordCount
RstProjCode.MoveLast
RstProjCode.MoveFirst
If RstProjCode.RecordCount = 0 Then
    GoTo Exit_writecode
End If
With RstProjCode
    
    mprojcodeID = !projcodeid
    Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                 (TblPT, nCODEHeadDwg, 1, 1, 1, 0) 'First dwg
    Dim varAttributes As Variant
    varAttributes = blockRefObj.GetAttributes
    For iattr_CODE = LBound(varAttributes) To UBound(varAttributes)
        Debug.Print varAttributes(iattr_CODE).TagString
'************ PROJ_NO ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XPROJ_NO" Then
            If IsNull(!PROJ_NO) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !PROJ_NO
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ DATE ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XDATE" Then
                varAttributes(iattr_CODE).textString = Date
        End If
'************ XPROJADDR ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XPROJADDR" Then
            If IsNull(!PROJ_ADDR) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !PROJ_ADDR
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XBLOCK ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XBLOCK" Then
            If IsNull(!Block) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !Block
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XLOT ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XLOT" Then
            If IsNull(!LOT) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !LOT
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XBIN ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XBIN" Then
            If IsNull(!BIN) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !BIN
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XCB ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XCB" Then
            If IsNull(!COMMBD) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !COMMBD
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XHT ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XHT" Then
            If IsNull(!BLDGHT_EX) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !BLDGHT_EX
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++
'++++++      START OF
'++++++      8/25/14 EDITED FOR MORE FIELDS
'++++++
'++++++
'++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++




'************ XHT_NW ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XHT_NW" Then
            If IsNull(!BLDGHT_NW) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !BLDGHT_NW
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XSTORY ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XSTORY" Then
            If IsNull(!STORY) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !STORY
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XSTORY_NW ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XSTORY_NW" Then
            If IsNull(!STORY_NW) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !STORY_NW
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XDU_EX ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XDU_EX" Then
            If IsNull(!DU_EX) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !DU_EX
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XDU_ADD ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XDU_ADD" Then
            If IsNull(!DU_ADD) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !DU_ADD
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XDU_NW ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XDU_NW" Then
            If IsNull(!DU_NW) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !DU_NW
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XDU_TOT ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XDU_TOT" Then
            If IsNull(!DU_TOTAL) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !DU_TOTAL
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XAREA_EX ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XAREA_EX" Then
            If IsNull(!AREA_EX) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !AREA_EX
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XAREA_ADD ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XAREA_ADD" Then
            If IsNull(!AREA_ADD) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !AREA_ADD
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XAREA_NW ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XAREA_NW" Then
            If IsNull(!AREA_NW) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !AREA_NW
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XAREA_TOT ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XAREA_TOT" Then
            If IsNull(!AREA_TOTAL) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !AREA_TOTAL
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++
'++++++      END OF
'++++++      8/25/14 EDITED FOR MORE FIELDS
'++++++
'++++++
'++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


'************ XZONE ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XZONE" Then
            If IsNull(!ZONE) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !ZONE
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XLOTSIZE ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XLOTSIZE" Then
            If IsNull(!LOTSIZE) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = Format(!LOTSIZE, "0,0") & " S.F. APPROX."
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XCOVERAGE_EX ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XCOVERAGE_EX" Then
            If IsNull(!COVERAGE_EXISTING) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !COVERAGE_EXISTING
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XCOVERAGE_NW ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XCOVERAGE_NW" Then
            If IsNull(!COVERAGE_NEW) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !COVERAGE_NEW
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XOVZONE ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XOVZONE" Then
            If IsNull(!OVZONE) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !OVZONE
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XMAP ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XMAP" Then
            If IsNull(!ZONEMAP) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !ZONEMAP
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
        
'************ XMUNI ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XMUNI" Then
            If IsNull(!MUNI) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !COMPANY
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
        If varAttributes(iattr_CODE).TagString = "XMUNISTREET" Then
            Debug.Print "+++++++++++++++++ Got XMUNIstreet ++++++" & varAttributes(iattr_MUNI).TagString
            If IsNull(!STREET) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !STREET
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
        If varAttributes(iattr_CODE).TagString = "XMUNICSZ" Then
            If IsNull(!csz) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !csz
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
        ''End With
    Next iattr_CODE
    TblPT(1) = TblPT(1) - nJumpHead
End With
RstProjCode.Close
'**********************************************************************************
'*    CODE routine
'**********************************************************************************
    
    strRefCode = "SELECT PROJCODE.PROJCODEID, CODE.CODENAME " _
    & "FROM CODE INNER JOIN (PROJCODE INNER JOIN CONNCODE ON " _
    & "PROJCODE.PROJCODEID = CONNCODE.PROJCODEID) ON CODE.CODEID = CONNCODE.CODEID " _
    & "WHERE (((PROJCODE.PROJCODEID)=" & mprojcodeID & "))"
     Set rstRefCode = DB.OpenRecordset(strRefCode)
     With rstRefCode
    nRefCode = rstRefCode.RecordCount
    .MoveLast
    .MoveFirst
nRefCode = rstRefCode.RecordCount
     For k = 1 To nRefCode
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                     (TblPT, nREFCODEDwg, 1, 1, 1, 0)
        Dim varAttributes2 As Variant
        varAttributes2 = blockRefObj.GetAttributes
        For iattr_refcode = LBound(varAttributes2) To UBound(varAttributes2)
        
            If varAttributes2(iattr_refcode).TagString = "XCODENAME" Then
                If IsNull(!CODENAME) = True Then
                    varAttributes2(iattr_refcode).textString = ""
                Else
                    varAttributes2(iattr_refcode).textString = !CODENAME
                    Debug.Print varAttributes(iattr_refcode).textString
                End If
            End If
        
        Next iattr_refcode
        'rstRefCode.Close
        TblPT(1) = TblPT(1) - nJumpREFCODEROW
        Debug.Print "K:" & k
    .MoveNext
    Next k
    End With
    rstRefCode.Close
    Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                    (TblPT, nDOBHDDwg, 1, 1, 1, 0)
    TblPT(1) = TblPT(1) - nJumpDOBHD
'**********************************************************************************
'**********************************************************************************
'*    BIS routine
'**********************************************************************************
    
    strBisRow = "SELECT PROJCODE.PROJCODEID, proj_dob.bis, proj_dob.projbisnote, " _
    & "proj_dob.dobjobid, proj_dob.projbistype " _
    & "FROM PROJCODE INNER JOIN proj_dob ON PROJCODE.PROJCODEID = proj_dob.projcodeid " _
    & "WHERE (((PROJCODE.PROJCODEID)=" & mprojcodeID & "))"

     Set rstBisRow = DB.OpenRecordset(strBisRow)
     If rstBisRow.RecordCount = 0 Then
     GoTo SKIP_BISROUTINE
     
     End If
     
     With rstBisRow
     .MoveLast
     .MoveFirst
     nBISrow = rstBisRow.RecordCount
     For l = 0 To nBISrow - 1
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                     (TblPT, nDOBRowDwg, 1, 1, 1, 0)
        Dim varAttributes3 As Variant
        varAttributes3 = blockRefObj.GetAttributes
        For iattr_BISrow = LBound(varAttributes3) To UBound(varAttributes3)
'************ XBO ***************************************************************
            'If varAttributes3(iattr_BISrow).TagString = "XBO" Then
                'If IsNull(!bisorder) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    'varAttributes3(iattr_BISrow).TextString = ""
                'Else
                    'varAttributes3(iattr_BISrow).TextString = !bisorder
                    
                'End If
            'End If
'************ XBIS ***************************************************************
            If varAttributes3(iattr_BISrow).TagString = "XBIS" Then
                If IsNull(!bis) = True Then
                    varAttributes3(iattr_BISrow).textString = ""
                Else
                    varAttributes3(iattr_BISrow).textString = !bis
                    
                End If
            End If
'************ XBISTYPE ***************************************************************
            If varAttributes3(iattr_BISrow).TagString = "XBISTYPE" Then
                If IsNull(!projbistype) = True Then
                    varAttributes3(iattr_BISrow).textString = ""
                Else
                    varAttributes3(iattr_BISrow).textString = !projbistype
                    
                End If
            End If
'************ XBISDESC1 ***************************************************************
            If varAttributes3(iattr_BISrow).TagString = "XBISDESC1" Then
                If IsNull(!projbisnote) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes3(iattr_BISrow).textString = ""
                Else
                    varAttributes3(iattr_BISrow).textString = SEPTEXT2(!projbisnote, 1, 3, 33)
                    
                End If
            End If
'************ XBISDESC2 ***************************************************************
            If varAttributes3(iattr_BISrow).TagString = "XBISDESC2" Then
                If IsNull(!projbisnote) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes3(iattr_BISrow).textString = ""
                Else
                    varAttributes3(iattr_BISrow).textString = SEPTEXT2(!projbisnote, 2, 3, 33)
                    
                End If
            End If
'************ XBISDESC1 ***************************************************************
            If varAttributes3(iattr_BISrow).TagString = "XBISDESC2" Then
                If IsNull(!projbisnote) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes3(iattr_BISrow).textString = ""
                Else
                    varAttributes3(iattr_BISrow).textString = SEPTEXT2(!projbisnote, 3, 3, 33)
                    
                End If
            End If



'************ XBO ***************************************************************
        Next iattr_BISrow
        
        TblPT(1) = TblPT(1) - nJumpBISRow
    .MoveNext
    Next l
    rstBisRow.Close
    End With
SKIP_BISROUTINE:
'**********************************************************************************
'**********************************************************************************
'*    UG/OCC routine
'**********************************************************************************
'**********************************************************************************
    
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                     (TblPT, nCODEBCHDDwg, 1, 1, 1, 0)
   TblPT(1) = TblPT(1) - nJumpOCCTYPEHD
'**********************************************************************************
    'strOCC = "SELECT CONNUSEGP.PROJCODEID, CONNUSEGP.PROJUSEGPNOTE, USEGP.USEGP, USEGP.USEGPNAME, USEGP.USEGPDESC " _
& "FROM USEGP INNER JOIN (PROJCODE INNER JOIN CONNUSEGP ON PROJCODE.PROJCODEID = CONNUSEGP.PROJCODEID) ON USEGP.USEGPID = CONNUSEGP.USEGPID " _
& "WHERE (((CONNUSEGP.PROJCODEID)=" & MPROJCODEID & "));"
strOCC = "SELECT CONNOCCGRP.CONNOCCGRPID, CONNOCCGRP.OCCGRPID, CONNOCCGRP.PROJCODEID, " _
& "CONNOCCGRP.PROJOCCGRPNOTE, OCCGRP.OCCGRP, OCCGRP.OCCGRPNAME, CONNOCCGRP.PROJ_OCCLOADTYPE, " _
& "CONNOCCGRP.PROJ_OCCLOAD, CONNOCCGRP.PROJ_OCCLOADUNIT, CONNOCCGRP.ProjOccLoadNote " _
& "FROM OCCGRP INNER JOIN CONNOCCGRP ON OCCGRP.OCCGRPID = CONNOCCGRP.OCCGRPID " _
& "WHERE (((CONNOCCGRP.PROJCODEID)=" & mprojcodeID & "));"

     Set rstOCC = DB.OpenRecordset(strOCC)
     With rstOCC
     .MoveLast
    .MoveFirst
     Do While Not rstOCC.EOF
         nOCCrow = rstOCC.RecordCount
        .MoveLast
        .MoveFirst
        If rstOCC.RecordCount = 0 Then
            GoTo NoOCCStep
        End If
        
         For m = 1 To nOCCrow
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                         (TblPT, nOCCTYPEROWDwg, 1, 1, 1, 0)
            Dim varAttributes4 As Variant
            varAttributes4 = blockRefObj.GetAttributes
            For iattr_OCC = LBound(varAttributes4) To UBound(varAttributes4)
    '************ XOCCTYPE ***************************************************************
                If varAttributes4(iattr_OCC).TagString = "XOCCTYPE" Then
                    If IsNull(!OCCGRP) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes4(iattr_OCC).textString = ""
                    Else
                        varAttributes4(iattr_OCC).textString = !OCCGRP
                        
                    End If
                End If
    '************ XOCCNAME ***************************************************************
                If varAttributes4(iattr_OCC).TagString = "XOCCNAME" Then
                    If IsNull(!OCCGRPNAME) = True Then
                        varAttributes4(iattr_OCC).textString = ""
                    Else
                        varAttributes4(iattr_OCC).textString = !OCCGRPNAME
                        
                    End If
                End If
    '************ XOCCRATE ***************************************************************
                If varAttributes4(iattr_OCC).TagString = "XOCCRATE" Then
                    If IsNull(!PROJ_OCCLOAD) = True Then
                        varAttributes4(iattr_OCC).textString = ""
                    Else
                        varAttributes4(iattr_OCC).textString = !PROJ_OCCLOAD
                        
                    End If
                End If
    '************ XOCCNOTE ***************************************************************
                If varAttributes4(iattr_OCC).TagString = "XOCCNOTE" Then
                    If IsNull(!PROJOCCGRPNOTE) = True Then
                        varAttributes4(iattr_OCC).textString = ""
                    Else
                        varAttributes4(iattr_OCC).textString = !PROJOCCGRPNOTE
                        
                    End If
                End If
    '************  ***************************************************************
            Next iattr_OCC
            
            TblPT(1) = TblPT(1) - nJumpOCCTYPEROW
        .MoveNext
        Next m
    Loop
    End With
'**********************************************************************************
'**********************************************************************************
'*    CONTYPE routine
'**********************************************************************************
'**********************************************************************************


strConType = "SELECT CONNCONTYPE.CONNCONTYPEID, CONNCONTYPE.CONTYPEID, " _
& "CONNCONTYPE.PROJCODEID, CONNCONTYPE.ProjCONTYPEnote, CONTYPE.CONTYPE, CONTYPE.CONTYPEDESC " _
& "FROM CONTYPE INNER JOIN CONNCONTYPE ON CONTYPE.CONTYPEID = CONNCONTYPE.CONTYPEID " _
& "WHERE (((CONNCONTYPE.PROJCODEID)=" & mprojcodeID & "));"


Set rstConType = DB.OpenRecordset(strConType)
With rstConType
    .MoveLast
    .MoveFirst
     Do While Not rstConType.EOF
       nCONTYPErow = rstConType.RecordCount
        Debug.Print nCONTYPErow
        If nCONTYPErow > 0 Then
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                              (TblPT, nCONTYPEHDDwg, 1, 1, 1, 0)
            TblPT(1) = TblPT(1) - nJumpCONTYPEHD
         
         End If
         
        .MoveLast
        .MoveFirst
        If rstConType.RecordCount = 0 Then
            GoTo NoCONTYPEStep
        End If
    '*************************************************************************************
    '************ 08282014 ***************************************************************
    '*************************************************************************************
        mCONNCONTYPEID = !CONNCONTYPEID
        MCONTYPEID = !CONTYPEID
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
        For O = 1 To nCONTYPErow
         
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                         (TblPT, nCONTYPEROWDwg, 1, 1, 1, 0)
            Dim varAttributes5 As Variant
            varAttributes5 = blockRefObj.GetAttributes
            For iattr_ConType = LBound(varAttributes5) To UBound(varAttributes5)
                'mCONNCONTYPEID = CONNCONTYPEID
    '************ XOCCTYPE ***************************************************************
                If varAttributes5(iattr_ConType).TagString = "XCONTYPE" Then
                    If IsNull(!CONTYPE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes5(iattr_ConType).textString = ""
                    Else
                        varAttributes5(iattr_ConType).textString = !CONTYPE
                        
                    End If
                End If
    '************ XOCCNAME ***************************************************************
                If varAttributes5(iattr_ConType).TagString = "XCONDESC" Then
                    If IsNull(!CONTYPEDESC) = True Then
                        varAttributes5(iattr_ConType).textString = ""
                    Else
                        varAttributes5(iattr_ConType).textString = !CONTYPEDESC
                        
                    End If
                End If
    '************ XOCCRATE ***************************************************************
                If varAttributes5(iattr_ConType).TagString = "XCONTYPENOTE" Then
                    If IsNull(!ProjCONTYPEnote) = True Then
                        varAttributes5(iattr_ConType).textString = ""
                    Else
                        varAttributes5(iattr_ConType).textString = !ProjCONTYPEnote
                        
                    End If
                 End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
            Next iattr_ConType
            
            TblPT(1) = TblPT(1) - nJumpOCCTYPEROW
                '###########################################################################
                '***************************************************************************
                '*   AREA/HT SUB ROUTINE
                '***************************************************************************
             If mCONNCONTYPEID <> "" Then
                'strAH = "SELECT CONNAREASTORY.CONNAREASTORYID, CONNAREASTORY.AREASTORYID, " _
                & "CONNAREASTORY.PROJCODEID, CONNAREASTORY.PROJAREASTORYNOTE, CONNAREASTORY.PROJADJAREA, " _
                & "CONNAREASTORY.PROJADJHT, AREASTORY.CONTYPE, AREASTORY.USEGPID, AREASTORY.HTFT, " _
                & "AREASTORY.STORY, AREASTORY.AREA, CONNAREASTORY.CONNCONTYPEID, USEGP.USEGP " _
                & "FROM (AREASTORY INNER JOIN CONNAREASTORY ON AREASTORY.ID = CONNAREASTORY.AREASTORYID) INNER JOIN USEGP ON AREASTORY.USEGPID = USEGP.USEGPID " _
                & "WHERE (((CONNAREASTORY.CONNCONTYPEID)=" & mCONNCONTYPEID & "));"
                
                'strAH = "SELECT CONNCONTYPE.CONNCONTYPEID,  OCCGRP.OCCGRP, CONNCONTYPE.CONTYPEID, CONNCONTYPE.PROJCODEID, CONNCONTYPE.ProjCONTYPEnote, CONNAREASTORY.PROJAREASTORYNOTE, " _
                & "CONNAREASTORY.PROJADJAREA, CONNAREASTORY.PROJADJHT, CONNAREASTORY.PROJADJSTORY, CONTYPE.CONTYPE, AREASTORY.HTFT, AREASTORY.STORY, AREASTORY.AREA " _
                & "FROM AREASTORY INNER JOIN ((CONTYPE INNER JOIN CONNCONTYPE ON CONTYPE.CONTYPEID = CONNCONTYPE.CONTYPEID) INNER JOIN CONNAREASTORY ON CONNCONTYPE.CONNCONTYPEID = CONNAREASTORY.CONNCONTYPEID) " _
                & "ON AREASTORY.ID = CONNAREASTORY.AREASTORYID " _
                & "WHERE (((CONNCONTYPE.CONNCONTYPEID)=" & mCONNCONTYPEID & "));"

                StrAH = "SELECT CONNCONTYPE.CONNCONTYPEID, OCCGRP.OCCGRP, CONNCONTYPE.CONTYPEID, CONNCONTYPE.PROJCODEID, CONNCONTYPE.ProjCONTYPEnote, CONNAREASTORY.PROJAREASTORYNOTE, " _
                & "CONNAREASTORY.PROJADJAREA, CONNAREASTORY.PROJADJHT, CONNAREASTORY.PROJADJSTORY, CONTYPE.CONTYPE, AREASTORY.HTFT, AREASTORY.STORY, AREASTORY.AREA " _
                & "FROM (CONTYPE INNER JOIN (AREASTORY INNER JOIN (CONNCONTYPE INNER JOIN CONNAREASTORY ON CONNCONTYPE.CONNCONTYPEID = CONNAREASTORY.CONNCONTYPEID) ON AREASTORY.ID = CONNAREASTORY.AREASTORYID) " _
                & "ON CONTYPE.CONTYPEID = CONNCONTYPE.CONTYPEID) INNER JOIN OCCGRP ON AREASTORY.occgrpID = OCCGRP.OCCGRPID " _
                & "WHERE (((CONNCONTYPE.projcodeID)=" & mprojcodeID & "));"
                
                '& "WHERE (((CONNCONTYPE.contypeID)=" & mcontypeID & "));"
                Set rstAH = DB.OpenRecordset(StrAH)
                With rstAH
                    .MoveLast
                    .MoveFirst
                     Do While Not rstAH.EOF
                        nAHrow = rstAH.RecordCount
                        If nAHrow > 0 Then
                            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                              (TblPT, nAHHDDwg, 1, 1, 1, 0)
                            TblPT(1) = TblPT(1) - nJumpAHHD
                         
                         End If
                         
                        .MoveLast
                        .MoveFirst
                        If rstAH.RecordCount = 0 Then
                            GoTo NoAHStep
                        End If
                        
                         For P = 1 To nAHrow
                            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                         (TblPT, nAHROWDwg, 1, 1, 1, 0)
                            Dim varAttributes6 As Variant
                            varAttributes6 = blockRefObj.GetAttributes
                            For iattr_AH = LBound(varAttributes6) To UBound(varAttributes6)
                    '************ XOCCTYPE ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XOCCTYPE" Then
                                    If IsNull(!OCCGRP) = True Then '++++++++++++++++++++++++++++++++++++++++++
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = !OCCGRP
                                        
                                    End If
                                End If
                    '************ XAHHT ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHHT" Then
                                    If IsNull(!HTFT) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = !HTFT
                                        
                                    End If
                                End If
                    '************ XAHS ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHS" Then
                                    If IsNull(!STORY) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = !STORY
                                        
                                    End If
                                 End If
                    '************ XAHA ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHA" Then
                                    If IsNull(!Area) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = Format(!Area, "0,0") & " SF"
                                        
                                    End If
                                 End If
                    '************ XAHPHT ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHPHT" Then
                                    If IsNull(!PROJADJHT) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = !PROJADJHT
                                        
                                    End If
                                 End If
                    '************ XAHPS ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHPS" Then
                                    If IsNull(!PROJADJSTORY) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = !PROJADJSTORY
                                        
                                    End If
                                 End If
                    '************ XAHPA ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHPA" Then
                                    If IsNull(!PROJADJAREA) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = Format(!PROJADJAREA, "0,0") & " SF"
                                        
                                    End If
                                 End If
                    '************ XAHNOTE ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHNOTE" Then
                                    If IsNull(!PROJAREASTORYNOTE) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = !PROJAREASTORYNOTE
                                        
                                    End If
                                 End If

                    '************************************************************************************
                           Next iattr_AH
                            TblPT(1) = TblPT(1) - nJumpAHROW

                           .MoveNext
                        Next P
                    Loop
                End With
                rstAH.Close
                End If
                '###########################################################################
                '*   END AREA/HT SUB ROUTINE
                '###########################################################################
                
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        .MoveNext
        Next O
    Loop
End With

'**********************************************************************************
'**********************************************************************************
'*    CONELEM routine
'**********************************************************************************
'**********************************************************************************
                
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++
'+++++ START OF CONELEM
'+++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'iattr_ConType, iattr_ExtElem0, iattr_ExtElem1

With rstConType 'with 1
    .MoveLast
    .MoveFirst
    Do While Not rstConType.EOF
    nCONTYPErow = rstConType.RecordCount
    'For s = 1 To nCONTYPErow
    If MCONTYPEID <> "" Then 'IF 0825
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                          (TblPT, nELEMHDDwg, 1, 1, 1, 0)
        TblPT(1) = TblPT(1) - nJumpCONELEMHD 'nJumpCONELEMHD
        'tblPT(1) = tblPT(1) - 36
        Dim varAttributes7 As Variant
        varAttributes7 = blockRefObj.GetAttributes
        
        '01###########################################################################
        For iattr_CONELEMHD = LBound(varAttributes7) To UBound(varAttributes7) 'Loop 01
            
            '************ XCONTYPE ***************************************************************
            If varAttributes7(iattr_CONELEMHD).TagString = "XCONTYPE" Then
                If IsNull(!CONTYPE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes7(iattr_CONELEMHD).textString = ""
                Else
                    varAttributes7(iattr_CONELEMHD).textString = !CONTYPE
                    
                End If
            End If
            '************ XCONDESC ***************************************************************
            If varAttributes7(iattr_CONELEMHD).TagString = "XCONDESC" Then
                If IsNull(!CONTYPEDESC) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes7(iattr_CONELEMHD).textString = ""
                Else
                    varAttributes7(iattr_CONELEMHD).textString = !CONTYPEDESC
                    
                End If
            End If
        Next iattr_CONELEMHD
            'nELEMHDDwg = "H:\drawfile\jvba\XCODE_CON_ELEM_HD01.dwg"
            'nELEM0Dwg = "H:\drawfile\jvba\XCODE_CON_ELEM0_ROW01.dwg"
            'nELEM1Dwg = "H:\drawfile\jvba\XCODE_CON_ELEM1_ROW01.dwg"
            'nJumpCONELEMHD = 36
            'nJumpCONELEM0ROW = 12
            'nJumpCONELEM1ROW = 12
            'strExtElem0 = "SELECT EXTELEM0.EXTELEM0ID, EXTELEM0.EXTELM0, EXTELEM0.EXTELM0LEVEL, " _
            & "EXTELEM0.EXTELEM0RATE, EXTELEM0.EXTELEM0DESC, EXTELEM0.CONTYPEID FROM EXTELEM0;"
            strExtElem0 = "SELECT EXTELEM0.EXTELEM0ID, EXTELEM0.EXTELM0, EXTELEM0.EXTELM0LEVEL, " _
            & "EXTELEM0.EXTELEM0RATE, EXTELEM0.EXTELEM0DESC, CONTYPE.CONTYPEID " _
            & "FROM EXTELEM0 INNER JOIN CONTYPE ON EXTELEM0.CONTYPEID = CONTYPE.CONTYPEID " _
            & "WHERE (((CONTYPE.CONTYPEID) =" & MCONTYPEID & "))" _
            & "ORDER BY EXTELEM0.EXTELM0LEVEL;"
            '###########################################################################
            Set rstExtElem0 = DB.OpenRecordset(strExtElem0)
            
            If rstExtElem0.RecordCount > 0 Then 'If 1.5
            With rstExtElem0 'with 2
                .MoveLast
                .MoveFirst
                nExtElem0row = rstExtElem0.RecordCount
                '02###########################################################################
                Do While Not rstExtElem0.EOF
                    'nExtElem0row = rstExtElem0.RecordCount
                    If nExtElem0row > 0 Then
                        'Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                          (tblPT, nELEM0Dwg, 1, 1, 1, 0)
                        'tblPT(1) = tblPT(1) - nJumpCONELEM0ROW
                        
                     End If
                     
                    .MoveLast
                    .MoveFirst
                    If rstExtElem0.RecordCount = 0 Then
                        GoTo NoExtElem0Step
                    End If
                        '03###########################################################################
                        For Q = 1 To nExtElem0row
                        Debug.Print "Q:" & Q
                            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                        (TblPT, nELEM0Dwg, 1, 1, 1, 0)
                            Dim varAttributes8 As Variant
                            mExtElem0ID = !EXTELEM0ID
                            varAttributes8 = blockRefObj.GetAttributes
                            '04###########################################################################
                            For iattr_ExtElem0 = LBound(varAttributes8) To UBound(varAttributes8)
                                
                                '************ XELEM0 ***************************************************************
                                If varAttributes8(iattr_ExtElem0).TagString = "XELEM0" Then
                                    If IsNull(!EXTELM0) = True Then '++++++++++++++++++++++++++++++++++++++++++
                                        varAttributes8(iattr_ExtElem0).textString = ""
                                    Else
                                        Debug.Print "stop:>>>>>" & !EXTELM0
                                        varAttributes8(iattr_ExtElem0).textString = !EXTELM0
                                        If !EXTELM0 = "INTERIOR ELEMENTS" Then
                                        Debug.Print "stop:>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & !EXTELM0
                                        End If
                                    End If
                                End If
                                '************ XELEM0 ***************************************************************
                                If varAttributes8(iattr_ExtElem0).TagString = "XELEM0NOTE" Then
                                    If IsNull(!EXTELEM0DESC) = True Then '++++++++++++++++++++++++++++++++++++++++++
                                        varAttributes8(iattr_ExtElem0).textString = ""
                                    Else
                                        varAttributes8(iattr_ExtElem0).textString = !EXTELEM0DESC
                                        
                                    End If
                                End If
                            
                            
                            Next iattr_ExtElem0
                            TblPT(1) = TblPT(1) - nJumpCONELEM0ROW
                       
                            strExtElem1 = "SELECT EXTELEM0.EXTELEM0ID, EXTELEM1.EXTELEM1ID, EXTELEM1.EXTELEM1, EXTELEM1.EXTELEM1LEVEL, " _
                            & "EXTELEM1.EXTELEM1RATEID, EXTELEM1.EXTELEM1DESC, FIRERATE.FIRERATE " _
                            & "FROM FIRERATE INNER JOIN (EXTELEM0 INNER JOIN EXTELEM1 ON EXTELEM0.EXTELEM0ID = EXTELEM1.EXTELEM0ID) ON FIRERATE.FIRERATEID = EXTELEM1.EXTELEM1RATEID " _
                            & "WHERE (((EXTELEM0.EXTELEM0ID) = " & mExtElem0ID & "))" _
                            & "ORDER BY EXTELEM1.EXTELEM1LEVEL;"
    
    
    
    'SELECT EXTELEM0.EXTELEM0ID, EXTELEM1.EXTELEM1ID, EXTELEM1.EXTELEM1, EXTELEM1.EXTELEM1LEVEL, EXTELEM1.EXTELEM1RATEID, EXTELEM1.EXTELEM1DESC, FIRERATE.FIRERATE
    'FROM FIRERATE INNER JOIN (EXTELEM0 INNER JOIN EXTELEM1 ON EXTELEM0.EXTELEM0ID = EXTELEM1.EXTELEM0ID) ON FIRERATE.FIRERATEID = EXTELEM1.EXTELEM1RATEID
    'WHERE (((EXTELEM0.EXTELEM0ID) = 58))
    'ORDER BY EXTELEM1.EXTELEM1LEVEL;
    
    
                            '###########################################################################
                            Set rstExtElem1 = DB.OpenRecordset(strExtElem1)
                                With rstExtElem1 'with 3
                                    .MoveLast
                                    .MoveFirst
                                    '05###########################################################################
                                     'Do While Not rstExtElem1.EOF
                                        nExtElem1row = rstExtElem1.RecordCount
                                        Debug.Print nExtElem1row
                                        If nExtElem1row > 0 Then
                                            'Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                                              (tblPT, nELEM0Dwg, 1, 1, 1, 0)
                                            'tblPT(1) = tblPT(1) - nJumpCONELEM1ROW
                                            
                                         End If
                                         
                                        .MoveLast
                                        .MoveFirst
                                        If rstExtElem1.RecordCount = 0 Then
                                            GoTo NoExtElem1Step
                                        End If
                                            '06###########################################################################
                                            For R = 1 To nExtElem1row
                                                Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                                            (TblPT, nELEM1Dwg, 1, 1, 1, 0)
                                                Dim varAttributes9 As Variant
                                                varAttributes9 = blockRefObj.GetAttributes
                                                '07###########################################################################
                                                mExtElem1ID = !extelem1id
                                                For iattr_ExtElem1 = LBound(varAttributes9) To UBound(varAttributes9)
                                                    
                                                    '************ XELEM1 ***************************************************************
                                                    If varAttributes9(iattr_ExtElem1).TagString = "XELEM1" Then
                                                        If IsNull(!EXTELeM1) = True Then '++++++++++++++++++++++++++++++++++++++++++
                                                            varAttributes9(iattr_ExtElem1).textString = ""
                                                        Else
                                                            varAttributes9(iattr_ExtElem1).textString = !EXTELeM1
                                                            
                                                        End If
                                                    End If
                                                    '************ XELEM1RATE ***************************************************************
                                                    If varAttributes9(iattr_ExtElem1).TagString = "XELEM1RATE" Then
                                                        If IsNull(!FIRERATE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                                                            varAttributes9(iattr_ExtElem1).textString = ""
                                                        Else
                                                            varAttributes9(iattr_ExtElem1).textString = !FIRERATE
                                                            
                                                        End If
                                                    End If
                                                    '************ XELEM1NOTE ***************************************************************
                                                    If varAttributes9(iattr_ExtElem1).TagString = "XELEM1NOTE" Then
                                                        If IsNull(!EXTELEM1DESC) = True Then '++++++++++++++++++++++++++++++++++++++++++
                                                            varAttributes9(iattr_ExtElem1).textString = ""
                                                        Else
                                                            varAttributes9(iattr_ExtElem1).textString = !EXTELEM1DESC
                                                            
                                                        End If
                                                    End If
                                                Next iattr_ExtElem1
                                                    TblPT(1) = TblPT(1) - nJumpCONELEM1ROW
                                                    .MoveNext
                                            Next R
    
                                    'Loop
                                End With 'with 3
                            'Next iattr_ExtElem0
                            'tblPT(1) = tblPT(1) - nJumpCONELEM0ROW
                            .MoveNext
                            '04###########################################################################
                        Next Q
                        '03###########################################################################
                        '.MoveNext
                Loop
                '02###########################################################################
            End With 'with 2
            End If 'if 1.5
    '###########################################################################
        'Next iattr_CONELEMHD     'LOOP1
            'tblPT(1) = tblPT(1) - nJumpCONTYPEHD
        .MoveNext
        '01###########################################################################
        
    End If 'IF 0825
    '.MoveNext

'Next s
Loop
End With 'with1

'**********************************************************************************
'**********************************************************************************
'*    CONTYPE_ELEM routine
'**********************************************************************************
'**********************************************************************************
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++
'+++++ END OF CONELEM
'+++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++





'**********************************************************************************
NoOCCStep:
    rstOCC.Close
'rstProjCode.Close
NoAHStep:
    'rstAH.Close
NoExtElem0Step:

NoExtElem1Step:

NoCONTYPEStep:
    rstConType.Close

Set DB = Nothing

Exit_writecode:
frmCode.Hide
End Sub
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################
'############################################################################


Public Sub writeCode_BACKUP()
'############################################################################
'############################################################################
'#          11/15/12 Working JIN
'#
'#
'#
'#
'#
'############################################################################
'############################################################################

Dim DB As Database
Dim RstProjCode, rstProjMuni, rstRefCode, rstBisRow, rstOCC As Recordset
Dim strProjCode, strProjMuni, strRefCode, strBisRow, strOCC As String
Dim nRefCode, nBISrow, nOCCrow As Integer    'ProjSet no.
Dim iattr_CODE, iattr_refcode, iattr_BISrow, iattr_OCC As Integer
'Dim rcArray As Variant
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
'Dim tmpArrayAss(0 To 1) As Double

'Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
'Dim height As Double
Dim mode As Long
'Dim prompt As String

Dim nCODEHeadDwg, nREFCODEDwg, nDOBHDDwg, nDOBRowDwg, nCODEBCHDDwg As Variant
Dim nJumpHead, nJumpREFCODEROW, nJumpDOBHD, nJumpBISRow As Integer
Dim blockRefObj As AcadBlockReference
Dim BlkObj As Object
Dim k, l As Integer

'frmWorkItem.ListBox1.Clear
CreateCNstyle
TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify

    nCODEHeadDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_HD01.dwg"
    nREFCODEDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_REFCODE_ROW01.dwg"
    nDOBHDDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_DOB_HD01.dwg"
    nDOBRowDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_BIS_ROW01.dwg"
    nCODEBCHDDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_BC_HD01.dwg"
    nOCCTYPEROWDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_OCC_ROW01.dwg"
    nCONTYPEHDDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_CONTYPE_HD01.dwg"
    nCONTYPEROWDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_CONTYPE_ROW01.dwg"
    nAHHDDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_AH_HD01.dwg"
    nAHROWDwg = "\\jwja-svr-10\drawfile\jvba\XCODE_AH_ROW01.dwg"
    
    nJumpHead = 204 '17 X 12
    nJumpREFCODEROW = 12 '1 X 12
    nJumpDOBHD = 18
    nJumpBISRow = 24 '2 X 12
    nJumpOCCTYPEHD = 42 '2 X 12
    nJumpOCCTYPEROW = 12 '2 X 12
    nJumpCONTYPEHD = 24
    nJumpCONTYPEROW = 12
    nJumpAHHD = 36
    nJumpAHROW = 12
    
    'InsertJBlock "H:\drawfile\jvba\XCODE_HD01.dwg", tblPT, "Model"

'tblPT(1) = tblPT(1) - nJumpHead

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\cl_org_outspec_02.mdb")
xProj_no = frmCode.seleproj
Debug.Print xAssID
strProjCode = "SELECT PROJCODE.PROJCODEID, PROJCODE.PROJ_NO, PROJCODE.OCCLOADID, " _
& "Project.PROJ_NAME, Project.PROJ_STNO, Project.PROJ_ADDR, Project.PROJ_CITY, " _
& "Project.PROJ_STATE, Project.PROJ_ZIP, Project.projdesc, Project.PROJ_TYPE, Project.TYPE, " _
& "Project.STATUS, Project.CID, PROJCODE.BLOCK, PROJCODE.LOT, PROJCODE.SECTION, " _
& "PROJCODE.MUNI, PROJCODE.ZONE, PROJCODE.LOTSIZE, PROJCODE.COVERAGE_EXISTING, " _
& "PROJCODE.COVERAGE_NEW, PROJCODE.COMMBD, PROJCODE.BIN, PROJCODE.STORY, PROJCODE.OVZONE, " _
& "PROJCODE.ZONEMAP, PROJCODE.BLDGHT_EX, PROJCODE.BLDGHT_NW, " _
& "Contact.COMPANY, Contact.STREET, Contact.CITY, Contact.STATE, Contact.ZIP, Contact.WEBADDR, Trim([city] & ', ' & [state] & ' ' & [zip]) AS csz " _
& "FROM (PROJCODE INNER JOIN Project ON PROJCODE.PROJ_NO = Project.PROJ_NO) INNER JOIN Contact ON PROJCODE.MUNI = Contact.CID " _
& "WHERE (((PROJCODE.PROJ_NO)='" & xProj_no & "'));"

Set RstProjCode = DB.OpenRecordset(strProjCode)

RstProjCode.MoveLast
RstProjCode.MoveFirst
If RstProjCode.RecordCount = 0 Then
    GoTo Exit_writecode
End If
With RstProjCode
    
    mprojcodeID = !projcodeid
    Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                 (TblPT, nCODEHeadDwg, 1, 1, 1, 0) 'First dwg
    Dim varAttributes As Variant
    varAttributes = blockRefObj.GetAttributes
    For iattr_CODE = LBound(varAttributes) To UBound(varAttributes)
        Debug.Print varAttributes(iattr_CODE).TagString
'************ PROJ_NO ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XPROJ_NO" Then
            If IsNull(!PROJ_NO) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !PROJ_NO
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ DATE ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XDATE" Then
                varAttributes(iattr_CODE).textString = Date
        End If
'************ XPROJADDR ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XPROJADDR" Then
            If IsNull(!PROJ_ADDR) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !PROJ_ADDR
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XBLOCK ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XBLOCK" Then
            If IsNull(!Block) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !Block
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XLOT ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XLOT" Then
            If IsNull(!LOT) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !LOT
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XBIN ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XBIN" Then
            If IsNull(!BIN) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !BIN
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XCB ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XCB" Then
            If IsNull(!COMMBD) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !COMMBD
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XHT ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XHT" Then
            If IsNull(!BLDGHT_EX) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !BLDGHT_EX
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XSTORY ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XSTORY" Then
            If IsNull(!STORY) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !STORY
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XZONE ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XZONE" Then
            If IsNull(!ZONE) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !ZONE
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XLOTSIZE ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XLOTSIZE" Then
            If IsNull(!LOTSIZE) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = Format(!LOTSIZE, "0,0") & " S.F. APPROX."
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XCOVERAGE_EX ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XCOVERAGE_EX" Then
            If IsNull(!COVERAGE_EXISTING) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !COVERAGE_EXISTING
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XCOVERAGE_NW ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XCOVERAGE_NW" Then
            If IsNull(!COVERAGE_NEW) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !COVERAGE_NEW
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XOVZONE ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XOVZONE" Then
            If IsNull(!OVZONE) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !OVZONE
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
'************ XMAP ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XMAP" Then
            If IsNull(!ZONEMAP) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !ZONEMAP
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
        
'************ XMUNI ***************************************************************
        If varAttributes(iattr_CODE).TagString = "XMUNI" Then
            If IsNull(!MUNI) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !COMPANY
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
        If varAttributes(iattr_CODE).TagString = "XMUNISTREET" Then
            Debug.Print "+++++++++++++++++ Got XMUNIstreet ++++++" & varAttributes(iattr_MUNI).TagString
            If IsNull(!STREET) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !STREET
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
        If varAttributes(iattr_CODE).TagString = "XMUNICSZ" Then
            If IsNull(!csz) = True Then
                varAttributes(iattr_CODE).textString = ""
            Else
                varAttributes(iattr_CODE).textString = !csz
                Debug.Print varAttributes(iattr_CODE).textString
            End If
        End If
        ''End With
    Next iattr_CODE
    TblPT(1) = TblPT(1) - nJumpHead
End With
RstProjCode.Close
'**********************************************************************************
'*    CODE routine
'**********************************************************************************
    
    strRefCode = "SELECT PROJCODE.PROJCODEID, CODE.CODENAME " _
    & "FROM CODE INNER JOIN (PROJCODE INNER JOIN CONNCODE ON " _
    & "PROJCODE.PROJCODEID = CONNCODE.PROJCODEID) ON CODE.CODEID = CONNCODE.CODEID " _
    & "WHERE (((PROJCODE.PROJCODEID)=" & mprojcodeID & "))"
     Set rstRefCode = DB.OpenRecordset(strRefCode)
     With rstRefCode
    nRefCode = rstRefCode.RecordCount
    .MoveLast
    .MoveFirst
nRefCode = rstRefCode.RecordCount
     For k = 1 To nRefCode
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                     (TblPT, nREFCODEDwg, 1, 1, 1, 0)
        Dim varAttributes2 As Variant
        varAttributes2 = blockRefObj.GetAttributes
        For iattr_refcode = LBound(varAttributes2) To UBound(varAttributes2)
        
            If varAttributes2(iattr_refcode).TagString = "XCODENAME" Then
                If IsNull(!CODENAME) = True Then
                    varAttributes2(iattr_refcode).textString = ""
                Else
                    varAttributes2(iattr_refcode).textString = !CODENAME
                    Debug.Print varAttributes(iattr_refcode).textString
                End If
            End If
        
        Next iattr_refcode
        'rstRefCode.Close
        TblPT(1) = TblPT(1) - nJumpREFCODEROW
        Debug.Print "K:" & k
    .MoveNext
    Next k
    End With
    rstRefCode.Close
    Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                    (TblPT, nDOBHDDwg, 1, 1, 1, 0)
    TblPT(1) = TblPT(1) - nJumpDOBHD
'**********************************************************************************
'**********************************************************************************
'*    BIS routine
'**********************************************************************************
    
    strBisRow = "SELECT PROJCODE.PROJCODEID, proj_dob.bis, proj_dob.projbisnote, " _
    & "proj_dob.dobjobid, proj_dob.projbistype " _
    & "FROM PROJCODE INNER JOIN proj_dob ON PROJCODE.PROJCODEID = proj_dob.projcodeid " _
    & "WHERE (((PROJCODE.PROJCODEID)=" & mprojcodeID & "))"

     Set rstBisRow = DB.OpenRecordset(strBisRow)
     If rstBisRow.RecordCount = 0 Then
     GoTo SKIP_BISROUTINE
     
     End If
     
     With rstBisRow
     .MoveLast
     .MoveFirst
     nBISrow = rstBisRow.RecordCount
     For l = 0 To nBISrow - 1
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                     (TblPT, nDOBRowDwg, 1, 1, 1, 0)
        Dim varAttributes3 As Variant
        varAttributes3 = blockRefObj.GetAttributes
        For iattr_BISrow = LBound(varAttributes3) To UBound(varAttributes3)
'************ XBO ***************************************************************
            'If varAttributes3(iattr_BISrow).TagString = "XBO" Then
                'If IsNull(!bisorder) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    'varAttributes3(iattr_BISrow).TextString = ""
                'Else
                    'varAttributes3(iattr_BISrow).TextString = !bisorder
                    
                'End If
            'End If
'************ XBIS ***************************************************************
            If varAttributes3(iattr_BISrow).TagString = "XBIS" Then
                If IsNull(!bis) = True Then
                    varAttributes3(iattr_BISrow).textString = ""
                Else
                    varAttributes3(iattr_BISrow).textString = !bis
                    
                End If
            End If
'************ XBISTYPE ***************************************************************
            If varAttributes3(iattr_BISrow).TagString = "XBISTYPE" Then
                If IsNull(!projbistype) = True Then
                    varAttributes3(iattr_BISrow).textString = ""
                Else
                    varAttributes3(iattr_BISrow).textString = !projbistype
                    
                End If
            End If
'************ XBISDESC1 ***************************************************************
            If varAttributes3(iattr_BISrow).TagString = "XBISDESC1" Then
                If IsNull(!projbisnote) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes3(iattr_BISrow).textString = ""
                Else
                    varAttributes3(iattr_BISrow).textString = SEPTEXT2(!projbisnote, 1, 3, 33)
                    
                End If
            End If
'************ XBISDESC2 ***************************************************************
            If varAttributes3(iattr_BISrow).TagString = "XBISDESC2" Then
                If IsNull(!projbisnote) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes3(iattr_BISrow).textString = ""
                Else
                    varAttributes3(iattr_BISrow).textString = SEPTEXT2(!projbisnote, 2, 3, 33)
                    
                End If
            End If
'************ XBISDESC1 ***************************************************************
            If varAttributes3(iattr_BISrow).TagString = "XBISDESC2" Then
                If IsNull(!projbisnote) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes3(iattr_BISrow).textString = ""
                Else
                    varAttributes3(iattr_BISrow).textString = SEPTEXT2(!projbisnote, 3, 3, 33)
                    
                End If
            End If



'************ XBO ***************************************************************
        Next iattr_BISrow
        
        TblPT(1) = TblPT(1) - nJumpBISRow
    .MoveNext
    Next l
    rstBisRow.Close
    End With
SKIP_BISROUTINE:
'**********************************************************************************
'**********************************************************************************
'*    UG/OCC routine
'**********************************************************************************
'**********************************************************************************
    
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                     (TblPT, nCODEBCHDDwg, 1, 1, 1, 0)
   TblPT(1) = TblPT(1) - nJumpOCCTYPEHD
'**********************************************************************************
    'strOCC = "SELECT CONNUSEGP.PROJCODEID, CONNUSEGP.PROJUSEGPNOTE, USEGP.USEGP, USEGP.USEGPNAME, USEGP.USEGPDESC " _
& "FROM USEGP INNER JOIN (PROJCODE INNER JOIN CONNUSEGP ON PROJCODE.PROJCODEID = CONNUSEGP.PROJCODEID) ON USEGP.USEGPID = CONNUSEGP.USEGPID " _
& "WHERE (((CONNUSEGP.PROJCODEID)=" & MPROJCODEID & "));"
strOCC = "SELECT CONNOCCGRP.CONNOCCGRPID, CONNOCCGRP.OCCGRPID, CONNOCCGRP.PROJCODEID, " _
& "CONNOCCGRP.PROJOCCGRPNOTE, OCCGRP.OCCGRP, OCCGRP.OCCGRPNAME, CONNOCCGRP.PROJ_OCCLOADTYPE, " _
& "CONNOCCGRP.PROJ_OCCLOAD, CONNOCCGRP.PROJ_OCCLOADUNIT, CONNOCCGRP.ProjOccLoadNote " _
& "FROM OCCGRP INNER JOIN CONNOCCGRP ON OCCGRP.OCCGRPID = CONNOCCGRP.OCCGRPID " _
& "WHERE (((CONNOCCGRP.PROJCODEID)=" & mprojcodeID & "));"

     Set rstOCC = DB.OpenRecordset(strOCC)
     With rstOCC
     .MoveLast
    .MoveFirst
     Do While Not rstOCC.EOF
         nOCCrow = rstOCC.RecordCount
        .MoveLast
        .MoveFirst
        If rstOCC.RecordCount = 0 Then
            GoTo NoOCCStep
        End If
        
         For m = 1 To nOCCrow
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                         (TblPT, nOCCTYPEROWDwg, 1, 1, 1, 0)
            Dim varAttributes4 As Variant
            varAttributes4 = blockRefObj.GetAttributes
            For iattr_OCC = LBound(varAttributes4) To UBound(varAttributes4)
    '************ XOCCTYPE ***************************************************************
                If varAttributes4(iattr_OCC).TagString = "XOCCTYPE" Then
                    If IsNull(!OCCGRP) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes4(iattr_OCC).textString = ""
                    Else
                        varAttributes4(iattr_OCC).textString = !OCCGRP
                        
                    End If
                End If
    '************ XOCCNAME ***************************************************************
                If varAttributes4(iattr_OCC).TagString = "XOCCNAME" Then
                    If IsNull(!OCCGRPNAME) = True Then
                        varAttributes4(iattr_OCC).textString = ""
                    Else
                        varAttributes4(iattr_OCC).textString = !OCCGRPNAME
                        
                    End If
                End If
    '************ XOCCRATE ***************************************************************
                If varAttributes4(iattr_OCC).TagString = "XOCCRATE" Then
                    If IsNull(!PROJ_OCCLOAD) = True Then
                        varAttributes4(iattr_OCC).textString = ""
                    Else
                        varAttributes4(iattr_OCC).textString = !PROJ_OCCLOAD
                        
                    End If
                End If
    '************ XOCCNOTE ***************************************************************
                If varAttributes4(iattr_OCC).TagString = "XOCCNOTE" Then
                    If IsNull(!PROJOCCGRPNOTE) = True Then
                        varAttributes4(iattr_OCC).textString = ""
                    Else
                        varAttributes4(iattr_OCC).textString = !PROJOCCGRPNOTE
                        
                    End If
                End If
    '************  ***************************************************************
            Next iattr_OCC
            
            TblPT(1) = TblPT(1) - nJumpOCCTYPEROW
        .MoveNext
        Next m
    Loop
    End With
'**********************************************************************************
'**********************************************************************************
'*    CONTYPE routine
'**********************************************************************************
'**********************************************************************************


strConType = "SELECT CONNCONTYPE.CONNCONTYPEID, CONNCONTYPE.CONTYPEID, " _
& "CONNCONTYPE.PROJCODEID, CONNCONTYPE.ProjCONTYPEnote, CONTYPE.CONTYPE, CONTYPE.CONTYPEDESC " _
& "FROM CONTYPE INNER JOIN CONNCONTYPE ON CONTYPE.CONTYPEID = CONNCONTYPE.CONTYPEID " _
& "WHERE (((CONNCONTYPE.PROJCODEID)=" & mprojcodeID & "));"


Set rstConType = DB.OpenRecordset(strConType)
With rstConType
    .MoveLast
    .MoveFirst
     Do While Not rstConType.EOF
        nCONTYPErow = rstConType.RecordCount
        If nCONTYPErow > 0 Then
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                              (TblPT, nCONTYPEHDDwg, 1, 1, 1, 0)
            TblPT(1) = TblPT(1) - nJumpCONTYPEHD
         
         End If
         
        .MoveLast
        .MoveFirst
        If rstConType.RecordCount = 0 Then
            GoTo NoCONTYPEStep
        End If
        mCONNCONTYPEID = !CONNCONTYPEID
        MCONTYPEID = !CONTYPEID
        For O = 1 To nCONTYPErow
         
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                         (TblPT, nCONTYPEROWDwg, 1, 1, 1, 0)
            Dim varAttributes5 As Variant
            varAttributes5 = blockRefObj.GetAttributes
            For iattr_ConType = LBound(varAttributes5) To UBound(varAttributes5)
                'mCONNCONTYPEID = CONNCONTYPEID
    '************ XOCCTYPE ***************************************************************
                If varAttributes5(iattr_ConType).TagString = "XCONTYPE" Then
                    If IsNull(!CONTYPE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes5(iattr_ConType).textString = ""
                    Else
                        varAttributes5(iattr_ConType).textString = !CONTYPE
                        
                    End If
                End If
    '************ XOCCNAME ***************************************************************
                If varAttributes5(iattr_ConType).TagString = "XCONDESC" Then
                    If IsNull(!CONTYPEDESC) = True Then
                        varAttributes5(iattr_ConType).textString = ""
                    Else
                        varAttributes5(iattr_ConType).textString = !CONTYPEDESC
                        
                    End If
                End If
    '************ XOCCRATE ***************************************************************
                If varAttributes5(iattr_ConType).TagString = "XCONTYPENOTE" Then
                    If IsNull(!ProjCONTYPEnote) = True Then
                        varAttributes5(iattr_ConType).textString = ""
                    Else
                        varAttributes5(iattr_ConType).textString = !ProjCONTYPEnote
                        
                    End If
                 End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
            Next iattr_ConType
            
            TblPT(1) = TblPT(1) - nJumpOCCTYPEROW
                '###########################################################################
                '***************************************************************************
                '*   AREA/HT SUB ROUTINE
                '***************************************************************************
             If mCONNCONTYPEID <> "" Then
                'strAH = "SELECT CONNAREASTORY.CONNAREASTORYID, CONNAREASTORY.AREASTORYID, " _
                & "CONNAREASTORY.PROJCODEID, CONNAREASTORY.PROJAREASTORYNOTE, CONNAREASTORY.PROJADJAREA, " _
                & "CONNAREASTORY.PROJADJHT, AREASTORY.CONTYPE, AREASTORY.USEGPID, AREASTORY.HTFT, " _
                & "AREASTORY.STORY, AREASTORY.AREA, CONNAREASTORY.CONNCONTYPEID, USEGP.USEGP " _
                & "FROM (AREASTORY INNER JOIN CONNAREASTORY ON AREASTORY.ID = CONNAREASTORY.AREASTORYID) INNER JOIN USEGP ON AREASTORY.USEGPID = USEGP.USEGPID " _
                & "WHERE (((CONNAREASTORY.CONNCONTYPEID)=" & mCONNCONTYPEID & "));"
                
                'strAH = "SELECT CONNCONTYPE.CONNCONTYPEID,  OCCGRP.OCCGRP, CONNCONTYPE.CONTYPEID, CONNCONTYPE.PROJCODEID, CONNCONTYPE.ProjCONTYPEnote, CONNAREASTORY.PROJAREASTORYNOTE, " _
                & "CONNAREASTORY.PROJADJAREA, CONNAREASTORY.PROJADJHT, CONNAREASTORY.PROJADJSTORY, CONTYPE.CONTYPE, AREASTORY.HTFT, AREASTORY.STORY, AREASTORY.AREA " _
                & "FROM AREASTORY INNER JOIN ((CONTYPE INNER JOIN CONNCONTYPE ON CONTYPE.CONTYPEID = CONNCONTYPE.CONTYPEID) INNER JOIN CONNAREASTORY ON CONNCONTYPE.CONNCONTYPEID = CONNAREASTORY.CONNCONTYPEID) " _
                & "ON AREASTORY.ID = CONNAREASTORY.AREASTORYID " _
                & "WHERE (((CONNCONTYPE.CONNCONTYPEID)=" & mCONNCONTYPEID & "));"

                StrAH = "SELECT CONNCONTYPE.CONNCONTYPEID, OCCGRP.OCCGRP, CONNCONTYPE.CONTYPEID, CONNCONTYPE.PROJCODEID, CONNCONTYPE.ProjCONTYPEnote, CONNAREASTORY.PROJAREASTORYNOTE, " _
                & "CONNAREASTORY.PROJADJAREA, CONNAREASTORY.PROJADJHT, CONNAREASTORY.PROJADJSTORY, CONTYPE.CONTYPE, AREASTORY.HTFT, AREASTORY.STORY, AREASTORY.AREA " _
                & "FROM (CONTYPE INNER JOIN (AREASTORY INNER JOIN (CONNCONTYPE INNER JOIN CONNAREASTORY ON CONNCONTYPE.CONNCONTYPEID = CONNAREASTORY.CONNCONTYPEID) ON AREASTORY.ID = CONNAREASTORY.AREASTORYID) " _
                & "ON CONTYPE.CONTYPEID = CONNCONTYPE.CONTYPEID) INNER JOIN OCCGRP ON AREASTORY.occgrpID = OCCGRP.OCCGRPID " _
                & "WHERE (((CONNCONTYPE.CONNCONTYPEID)=" & mCONNCONTYPEID & "));"
                
                
                Set rstAH = DB.OpenRecordset(StrAH)
                With rstAH
                    .MoveLast
                    .MoveFirst
                     Do While Not rstAH.EOF
                        nAHrow = rstAH.RecordCount
                        If nAHrow > 0 Then
                            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                              (TblPT, nAHHDDwg, 1, 1, 1, 0)
                            TblPT(1) = TblPT(1) - nJumpAHHD
                         
                         End If
                         
                        .MoveLast
                        .MoveFirst
                        If rstAH.RecordCount = 0 Then
                            GoTo NoAHStep
                        End If
                        
                         For P = 1 To nAHrow
                            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                         (TblPT, nAHROWDwg, 1, 1, 1, 0)
                            Dim varAttributes6 As Variant
                            varAttributes6 = blockRefObj.GetAttributes
                            For iattr_AH = LBound(varAttributes6) To UBound(varAttributes6)
                    '************ XOCCTYPE ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XOCCTYPE" Then
                                    If IsNull(!OCCGRP) = True Then '++++++++++++++++++++++++++++++++++++++++++
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = !OCCGRP
                                        
                                    End If
                                End If
                    '************ XAHHT ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHHT" Then
                                    If IsNull(!HTFT) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = !HTFT
                                        
                                    End If
                                End If
                    '************ XAHS ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHS" Then
                                    If IsNull(!STORY) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = !STORY
                                        
                                    End If
                                 End If
                    '************ XAHA ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHA" Then
                                    If IsNull(!Area) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = Format(!Area, "0,0") & " SF"
                                        
                                    End If
                                 End If
                    '************ XAHPHT ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHPHT" Then
                                    If IsNull(!PROJADJHT) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = !PROJADJHT
                                        
                                    End If
                                 End If
                    '************ XAHPS ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHPS" Then
                                    If IsNull(!PROJADJSTORY) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = !PROJADJSTORY
                                        
                                    End If
                                 End If
                    '************ XAHPA ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHPA" Then
                                    If IsNull(!PROJADJAREA) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = Format(!PROJADJAREA, "0,0") & " SF"
                                        
                                    End If
                                 End If
                    '************ XAHNOTE ***************************************************************
                                If varAttributes6(iattr_AH).TagString = "XAHNOTE" Then
                                    If IsNull(!PROJAREASTORYNOTE) = True Then
                                        varAttributes6(iattr_AH).textString = ""
                                    Else
                                        varAttributes6(iattr_AH).textString = !PROJAREASTORYNOTE
                                        
                                    End If
                                 End If

                    '************************************************************************************
                           Next iattr_AH
                            TblPT(1) = TblPT(1) - nJumpAHROW

                           .MoveNext
                        Next P
                    Loop
                End With
                rstAH.Close
                End If
                '###########################################################################
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        .MoveNext
        Next O
    Loop
End With

'**********************************************************************************
'**********************************************************************************
'*    CONELEM routine
'**********************************************************************************
'**********************************************************************************





'**********************************************************************************
NoOCCStep:
    rstOCC.Close
'rstProjCode.Close
NoAHStep:
    'rstAH.Close
NoCONTYPEStep:
    rstConType.Close

Set DB = Nothing

Exit_writecode:
frmCode.Hide
End Sub



