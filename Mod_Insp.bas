Attribute VB_Name = "Mod_Insp"
Public Sub ShowInsp()
frmInsp.show
End Sub
Public Sub writeInsp()


'############################################################################
'############################################################################
'#         02/13 Working JIN
'#
'#
'#
'#
'#
'############################################################################
'############################################################################

Dim DB As Database
Dim rstProjInsp As Recordset
Dim strProjInsp As String
Dim nInspRow As Integer    'ProjSet no.
Dim iattr_Insp As Integer
'Dim rcArray As Variant
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
'Dim tmpArrayAss(0 To 1) As Double

'Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
'Dim height As Double
'Dim mode As Long
'Dim prompt As String

Dim nInspHeadDwg, nInspRowDwg As Variant
Dim nJumpHead, nJumpRow As Integer
Dim blockRefObj As AcadBlockReference
Dim BlkObj As Object
Dim k, l As Integer


CreateCNstyle
TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify

If frmInsp.simplelist = True Then
    nInspHeadDwg = "\\jwja-svr-10\drawfile\jvba\XINSPHEAD01B.dwg"
    nInspRowDwg = "\\jwja-svr-10\drawfile\jvba\XINSPROW01B.dwg"
        
    nJumpHead = 84 '11.5 X 12/ 7x12
    nJumpRow = 24 '3.5 X 12 / 2 x 12
    'InsertJBlock "H:\drawfile\jvba\X_WORKITEM_HEAD02.dwg", tblPT, "Model"
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\XINSPHEAD01B.dwg", TblPT, "Model"
    TblPT(1) = TblPT(1) - nJumpHead

Else
    nInspHeadDwg = "\\jwja-svr-10\drawfile\jvba\X_INSP_HEAD_04.dwg"
    nInspRowDwg = "\\jwja-svr-10\drawfile\jvba\X_INSP_ROW_04.dwg"
        
    nJumpHead = 138 '11.5 X 12
    nJumpRow = 42 '3.5 X 12
    'InsertJBlock "H:\drawfile\jvba\X_WORKITEM_HEAD02.dwg", tblPT, "Model"
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\X_INSP_HEAD_04.dwg", TblPT, "Model"
    TblPT(1) = TblPT(1) - nJumpHead
End If

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_DOB_INSPECTIONS\X-dob_inspections_01.mdb")

xInspId = frmInsp.seleproj
Debug.Print xInspId
'strProjCode = "SELECT PROJCODE.PROJCODEID, PROJCODE.PROJ_NO, PROJCODE.OCCLOADID, " _
& "Project.PROJ_NAME, Project.PROJ_STNO, Project.PROJ_ADDR, Project.PROJ_CITY, " _
& "Project.PROJ_STATE, Project.PROJ_ZIP, Project.projdesc, Project.PROJ_TYPE, Project.TYPE, " _
& "Project.STATUS, Project.CID, PROJCODE.BLOCK, PROJCODE.LOT, PROJCODE.SECTION, " _
& "PROJCODE.MUNI, PROJCODE.ZONE, PROJCODE.LOTSIZE, PROJCODE.COVERAGE_EXISTING, " _
& "PROJCODE.COVERAGE_NEW, PROJCODE.COMMBD, PROJCODE.BIN, PROJCODE.STORY, PROJCODE.OVZONE, " _
& "PROJCODE.ZONEMAP, PROJCODE.BLDGHT_EX, PROJCODE.BLDGHT_NW, " _
& "Contact.COMPANY, Contact.STREET, Contact.CITY, Contact.STATE, Contact.ZIP, Contact.WEBADDR, Trim([city] & ', ' & [state] & ' ' & [zip]) AS csz " _
& "FROM (PROJCODE INNER JOIN Project ON PROJCODE.PROJ_NO = Project.PROJ_NO) INNER JOIN Contact ON PROJCODE.MUNI = Contact.CID " _
& "WHERE (((PROJCODE.PROJ_NO)='" & xProj_no & "'));"

strProjInsp = "SELECT PROJINSPITEM.proj_no, PROJINSPITEM.projinspid, PROJINSPITEM.INPUT, PROJINSPITEM.INSPNOTE, " _
& "INSPECTION.INSPTYPE, INSPECTION.CODENO, INSPECTION.INSPNAME, INSPECTION.LEADDAYS, PROJINSPITEM.INSTRUCTION " _
& "FROM INSPECTION INNER JOIN PROJINSPITEM ON INSPECTION.INSPID = PROJINSPITEM.INSPID " _
& "WHERE (((PROJINSPITEM.projinspid) = " & xInspId & ")) " _
& "ORDER BY INSPECTION.INSPTYPE, INSPECTION.CODENO;"


Set rstProjInsp = DB.OpenRecordset(strProjInsp)
rstProjInsp.MoveFirst
rstProjInsp.MoveLast
rstProjInsp.MoveFirst
If rstProjInsp.RecordCount = 0 Then
    GoTo Exit_writeINSP
End If
nInspRow = rstProjInsp.RecordCount

With rstProjInsp
    MPROJinspID = !PROJinspID
    For l = 0 To nInspRow - 1
    Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
             (TblPT, nInspRowDwg, 1, 1, 1, 0) 'First dwg
             
    Dim varAttributes As Variant
    varAttributes = blockRefObj.GetAttributes
    
    For iattr_Insp = LBound(varAttributes) To UBound(varAttributes)

        Debug.Print varAttributes(iattr_Insp).TagString
'************ XTYPE ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XTYPE" Then
            If IsNull(!INSPTYPE) = True Then
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = !INSPTYPE
                Debug.Print varAttributes(iattr_Insp).textString
            End If
        End If
'************ XCODE1 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XCODE1" Then
            If IsNull(!Codeno) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!Codeno, 1, 2, 16)
                
            End If
        End If
'************ XCODE2 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XCODE2" Then
            If IsNull(!Codeno) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!Codeno, 2, 2, 16)
                
            End If
        End If
'************ XINSP1 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINSP1" Then
            If IsNull(!INSPNAME) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPNAME, 1, 3, 16)
            End If
        End If
'************ XINSP2 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINSP2" Then
            If IsNull(!INSPNAME) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPNAME, 2, 3, 16)
            End If
        End If
'************ XINSP3 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINSP3" Then
            If IsNull(!INSPNAME) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPNAME, 3, 3, 16)
            End If
        End If
'************ XINST1 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST1" Then
            If IsNull(!INSTRUCTION) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSTRUCTION, 1, 6, 64)
            End If
        End If
'************ XINST2 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST2" Then
            If IsNull(!INSTRUCTION) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSTRUCTION, 2, 6, 64)
            End If
        End If
'************ XINST3 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST3" Then
            If IsNull(!INSTRUCTION) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSTRUCTION, 3, 6, 64)
            End If
        End If
'************ XINST4 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST4" Then
            If IsNull(!INSTRUCTION) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSTRUCTION, 4, 6, 64)
            End If
        End If
'************ XINST5 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST5" Then
            If IsNull(!INSTRUCTION) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSTRUCTION, 5, 6, 64)
            End If
        End If
'************ XINST6 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST6" Then
            If IsNull(!INSTRUCTION) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSTRUCTION, 6, 6, 64)
            End If
        End If
'************ XLEAD ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XLEAD" Then
            If IsNull(!LEADDAYS) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = !LEADDAYS
            End If
        End If
'************ XNOTE1 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XNOTE1" Then
            If IsNull(!INSPNOTE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPNOTE, 1, 4, 36)
            End If
        End If
'************ XNOTE2 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XNOTE2" Then
            If IsNull(!INSPNOTE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPNOTE, 2, 4, 36)
            End If
        End If
'************ XNOTE3 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XNOTE3" Then
            If IsNull(!INSPNOTE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPNOTE, 3, 4, 36)
            End If
        End If
'************ XNOTE4 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XNOTE4" Then
            If IsNull(!INSPNOTE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPNOTE, 4, 4, 36)
            End If
        End If

'************ END TEXT INSERT LOOP ***************************************************************

    Next iattr_Insp
     TblPT(1) = TblPT(1) - nJumpRow
    .MoveNext
        
    Next l
   rstProjInsp.Close
End With

Set DB = Nothing

Exit_writeINSP:
frmInsp.Hide
End Sub

