Attribute VB_Name = "Mod_Energy"
Public Sub writeEnergyInsp()


'############################################################################
'############################################################################
'#         01/19/17
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

nInspHeadDwg = "\\jwja-svr-10\drawfile\jvba\X_INSP_HEAD_10.dwg"
nInspRowDwg = "\\jwja-svr-10\drawfile\jvba\X_INSP_ROW_10.dwg"
    
nJumpHead = 138 '11.5 X 12
'nJumpRow = 42 '3.5 X 12
nJumpRow = 60 '5 X 12
'InsertJBlock "H:\drawfile\jvba\X_WORKITEM_HEAD02.dwg", tblPT, "Model"
InsertJBlock "\\jwja-svr-10\drawfile\jvba\X_INSP_HEAD_10.dwg", TblPT, "Model"
TblPT(1) = TblPT(1) - nJumpHead

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_DOB_INSPECTIONS\X-dob_inspections_02.mdb")

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

'******* ENERGY Query "ENG-QRY-10"
strProjInsp = "SELECT PROJINSPITEM.proj_no, PROJINSPITEM.projinspid, PROJINSPITEM.INPUT, PROJINSPITEM.INSPNOTE, " _
& "INSPECTION.INSPTYPE, INSPECTION.CODENO, INSPECTION.INSPNAME, INSPECTION.LEADDAYS, PROJINSPITEM.INSTRUCTION, " _
& "INSPECTION.INSPCRITERIOR, INSPECTION.INSPFREQ, INSPECTION.INSPREF, INSPECTION.INSPCITATION, INSPECTION.ITEM_INSTRUCTION " _
& "FROM INSPECTION INNER JOIN PROJINSPITEM ON INSPECTION.INSPID = PROJINSPITEM.INSPID " _
& "WHERE (((PROJINSPITEM.projinspid) = " & xInspId & ") AND ((INSPECTION.INSPTYPE)= 'ENERGY'))" _
& "ORDER BY INSPECTION.INSPTYPE, INSPECTION.CODENO, INSPECTION.INSPID;"


Set rstProjInsp = DB.OpenRecordset(strProjInsp)
rstProjInsp.MoveFirst
rstProjInsp.MoveLast
rstProjInsp.MoveFirst
If rstProjInsp.RecordCount = 0 Then
    GoTo Exit_writeENERGYINSP
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
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPNAME, 1, 4, 15)
            End If
        End If
'************ XINSP2 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINSP2" Then
            If IsNull(!INSPNAME) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPNAME, 2, 4, 15)
            End If
        End If
'************ XINSP3 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINSP3" Then
            If IsNull(!INSPNAME) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPNAME, 3, 4, 15)
            End If
        End If
'************ XINSP4 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINSP4" Then
            If IsNull(!INSPNAME) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPNAME, 4, 4, 15)
            End If
        End If

'************ XINST1 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST1" Then
            If IsNull(!INSPCRITERIOR) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCRITERIOR, 1, 8, 48)
            End If
        End If
'************ XINST2 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST2" Then
            If IsNull(!INSPCRITERIOR) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCRITERIOR, 2, 8, 48)
            End If
        End If
'************ XINST3 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST3" Then
            If IsNull(!INSPCRITERIOR) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCRITERIOR, 3, 8, 48)
            End If
        End If
'************ XINST4 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST4" Then
            If IsNull(!INSPCRITERIOR) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCRITERIOR, 4, 8, 48)
            End If
        End If
'************ XINST5 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST5" Then
            If IsNull(!INSPCRITERIOR) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCRITERIOR, 5, 8, 48)
            End If
        End If
'************ XINST6 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST6" Then
            If IsNull(!INSPCRITERIOR) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCRITERIOR, 6, 8, 48)
            End If
        End If
'************ XINST7 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST7" Then
            If IsNull(!INSPCRITERIOR) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCRITERIOR, 7, 8, 48)
            End If
        End If
'************ XINST8 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XINST8" Then
            If IsNull(!INSPCRITERIOR) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCRITERIOR, 8, 8, 48)
            End If
        End If

'************ XFREQ1 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XFREQ1" Then
            If IsNull(!INSPFREQ) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPFREQ, 1, 6, 18)
            End If
        End If
'************ XFREQ2 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XFREQ2" Then
            If IsNull(!INSPFREQ) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPFREQ, 2, 6, 18)
            End If
        End If
'************ XFREQ3 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XFREQ3" Then
            If IsNull(!INSPFREQ) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPFREQ, 3, 6, 18)
            End If
        End If
'************ XFREQ4 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XFREQ4" Then
            If IsNull(!INSPFREQ) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPFREQ, 4, 6, 18)
            End If
        End If
'************ XFREQ5 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XFREQ5" Then
            If IsNull(!INSPFREQ) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPFREQ, 5, 6, 18)
            End If
        End If
'************ XFREQ6 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XFREQ6" Then
            If IsNull(!INSPFREQ) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPFREQ, 6, 6, 18)
            End If
        End If


'************ XNOTE1 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XNOTE1" Then
            If IsNull(!INSPREF) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPREF, 1, 6, 18)
            End If
        End If
'************ XNOTE2 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XNOTE2" Then
            If IsNull(!INSPREF) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPREF, 2, 6, 18)
            End If
        End If
'************ XNOTE3 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XNOTE3" Then
            If IsNull(!INSPREF) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPREF, 3, 6, 18)
            End If
        End If
'************ XNOTE4 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XNOTE4" Then
            If IsNull(!INSPREF) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPREF, 4, 6, 18)
            End If
        End If
'************ XNOTE5 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XNOTE5" Then
            If IsNull(!INSPREF) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPREF, 5, 6, 18)
            End If
        End If
'************ XNOTE6 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "XNOTE6" Then
            If IsNull(!INSPREF) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPREF, 6, 6, 18)
            End If
        End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'************ XECC1 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "ECC1" Then
            If IsNull(!INSPCITATION) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCITATION, 1, 6, 18)
            End If
        End If
'************ XNOTE2 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "ECC2" Then
            If IsNull(!INSPCITATION) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCITATION, 2, 6, 18)
            End If
        End If
'************ XNOTE3 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "ECC3" Then
            If IsNull(!INSPCITATION) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCITATION, 3, 6, 18)
            End If
        End If
'************ XNOTE4 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "ECC4" Then
            If IsNull(!INSPCITATION) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCITATION, 4, 6, 18)
            End If
        End If
'************ XNOTE5 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "ECC5" Then
            If IsNull(!INSPCITATION) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCITATION, 5, 6, 18)
            End If
        End If
'************ XNOTE6 ***************************************************************
        If varAttributes(iattr_Insp).TagString = "ECC6" Then
            If IsNull(!INSPCITATION) = True Then '++++++++++++++++++++++++++++++++++++++++++
                varAttributes(iattr_Insp).textString = ""
            Else
                varAttributes(iattr_Insp).textString = SEPTEXT2(!INSPCITATION, 6, 6, 18)
            End If
        End If

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'************ END TEXT INSERT LOOP ***************************************************************

    Next iattr_Insp
     TblPT(1) = TblPT(1) - nJumpRow
    .MoveNext
        
    Next l
   rstProjInsp.Close
End With

Set DB = Nothing

Exit_writeENERGYINSP:
frmInsp.Hide
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub writeEnergyTabular()


'############################################################################
'############################################################################
'#         01/30/17
'#
'#
'#
'#
'#
'############################################################################
'############################################################################

Dim DB As Database
Dim rstProjECCItem As Recordset
Dim strProjECCItem As String
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
Dim blockRefObj, blockRefObj1 As AcadBlockReference
Dim BlkObj As Object
Dim k, l As Integer


CreateCNstyle
TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify

'nInspHeadDwg = "H:\drawfile\jvba\X_INSP_HEAD_10.dwg"
'nInspRowDwg = "H:\drawfile\jvba\X_INSP_ROW_10.dwg"
nInspHeadDwg = "\\jwja-svr-10\drawfile\jvba\XENGYHD05.dwg"
nInspRowDwg = "\\jwja-svr-10\drawfile\jvba\XENGYITEM05.dwg"
    
    
    
nJumpHead = 96 '11.5 X 12 ; 8 X 12
'nJumpRow = 42 '3.5 X 12
nJumpRow = 48 '5 X 12 ; 3 X 12 ; 4 X 12
'InsertJBlock "H:\drawfile\jvba\X_WORKITEM_HEAD02.dwg", tblPT, "Model"

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_DOB_INSPECTIONS\X-dob_inspections_02.mdb")

xProjECCId = frmInsp.seleProjECC
Debug.Print xProjECCId
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

'strProjInsp = "SELECT PROJINSPITEM.proj_no, PROJINSPITEM.projinspid, PROJINSPITEM.INPUT, PROJINSPITEM.INSPNOTE, " _
& "INSPECTION.INSPTYPE, INSPECTION.CODENO, INSPECTION.INSPNAME, INSPECTION.LEADDAYS, PROJINSPITEM.INSTRUCTION, " _
& "INSPECTION.INSPCRITERIOR, INSPECTION.INSPFREQ, INSPECTION.INSPREF, INSPECTION.INSPCITATION, INSPECTION.ITEM_INSTRUCTION " _
& "FROM INSPECTION INNER JOIN PROJINSPITEM ON INSPECTION.INSPID = PROJINSPITEM.INSPID " _
& "WHERE (((PROJINSPITEM.projinspid) = " & xInspId & ") AND ((INSPECTION.INSPTYPE)= 'ENERGY'))" _
& "ORDER BY INSPECTION.INSPTYPE, INSPECTION.CODENO, INSPECTION.INSPID;"

strProjECCItem = "SELECT PROJECC.CLIMATEZONE, PROJECC.CODE, PROJECC.ENERGYWORKSCOPE, PROJECC.PROJ_NO, PROJECCITEM.projeccitemid, PROJECCITEM.projeccid, " _
& "PROJECCITEM.ecctabid, PROJECCITEM.projeccdesc, PROJECCITEM.projeccpropvalue, " _
& "PROJECCITEM.projcodevalue, PROJECCITEM.projsuppdoc, PROJECCITEM.projnote, " _
& "PROJECCITEM.input, PROJECCITEM.projcodeid, PROJECCITEM.projcitation, PROJECCITEM.projprovision " _
& "FROM PROJECC INNER JOIN PROJECCitem ON PROJECC.projeccid = PROJECCITEM.projeccid  " _
& "WHERE (((PROJECCITEM.projeccid)= " & xProjECCId & ")); "

Set rstProjECCItem = DB.OpenRecordset(strProjECCItem)
rstProjECCItem.MoveFirst
rstProjECCItem.MoveLast
rstProjECCItem.MoveFirst
If rstProjECCItem.RecordCount = 0 Then
    GoTo Exit_writeENERGYTABULAR
End If
nInspRow = rstProjECCItem.RecordCount

With rstProjECCItem
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'InsertJBlock "Z:\jvba\XENGYHD04.dwg", tblPT, "Model"
    Set blockRefObj1 = ThisDrawing.ModelSpace.InsertBlock _
                 (TblPT, nInspHeadDwg, 1, 1, 1, 0) 'First dwg
    TblPT(1) = TblPT(1) - nJumpHead
    Dim varAttributes1 As Variant
    'Dim EX As Integer
    varAttributes1 = blockRefObj1.GetAttributes
    'Debug.Print !PROJECC.CODE
    For iattr_Insp = LBound(varAttributes1) To UBound(varAttributes1)
    '************ PROJ_NO ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "XPROJ_NO" Then
                If IsNull(!PROJ_NO) = True Then
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = !PROJ_NO
                    Debug.Print varAttributes1(iattr_Insp).textString
                End If
            End If
    '************ XDATE ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "XDATE" Then
                'If IsNull(!CODENO) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    'varAttributes(iattr_Insp).TextString = ""
                'Else
                    varAttributes1(iattr_Insp).textString = Date
                    
                'End If
            End If
    '************ CLIMATE ZONE ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "CLIMATEZONE" Then
                If IsNull(!climatezone) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = !climatezone
                    
                End If
            End If
    '************ CODE ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "CODE" Then
                If IsNull(!CODE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = !CODE
                    
                End If
            End If
    '************ SCOPE01 ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "SCOPE01" Then
                If IsNull(!ENERGYWORKSCOPE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = SEPTEXT2(!ENERGYWORKSCOPE, 1, 2, 40)
                    
                End If
            End If
    '************ SCOPE02 ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "SCOPE02" Then
                If IsNull(!ENERGYWORKSCOPE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = SEPTEXT2(!ENERGYWORKSCOPE, 2, 2, 40)
                    
                End If
            End If
    Next iattr_Insp
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    'MPROJinspID = !PROJinspID
    For l = 0 To nInspRow - 1
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                 (TblPT, nInspRowDwg, 1, 1, 1, 0) 'First dwg
                 
        Dim varAttributes As Variant
        Dim EX As Integer
        varAttributes = blockRefObj.GetAttributes
    
        For iattr_Insp = LBound(varAttributes) To UBound(varAttributes)
    
            Debug.Print varAttributes(iattr_Insp).TagString
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            '************ CITATATION ***************************************************************
            For EX = 1 To 5
                Debug.Print Trim("ENGY_ITEM_CIT_0" & EX)
                If varAttributes(iattr_Insp).TagString = Trim("ENGY_ITEM_CIT_0" & EX) Then
                    If IsNull(!projcitation) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes(iattr_Insp).textString = ""
                    Else
                        varAttributes(iattr_Insp).textString = SEPTEXT2(!projcitation, EX, 5, 12)
                    End If
                End If
            Next EX
            '************ PROVISION projprovision***************************************************************
            For EX = 1 To 5
                Debug.Print Trim("ENGY_ITEM_PROV_0" & EX)
                If varAttributes(iattr_Insp).TagString = Trim("ENGY_ITEM_PROV_0" & EX) Then
                    If IsNull(!projprovision) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes(iattr_Insp).textString = ""
                    Else
                        varAttributes(iattr_Insp).textString = SEPTEXT2(!projprovision, EX, 5, 22)
                    End If
                End If
            Next EX
            
            '************ DESCRIPTION projeccdesc***************************************************************
            For EX = 1 To 5
                Debug.Print Trim("ENGY_ITEM_DESC_0" & EX)
                If varAttributes(iattr_Insp).TagString = Trim("ENGY_ITEM_DESC_0" & EX) Then
                    If IsNull(!projeccdesc) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes(iattr_Insp).textString = ""
                    Else
                        varAttributes(iattr_Insp).textString = SEPTEXT2(!projeccdesc, EX, 5, 22)
                    End If
                End If
            Next EX
            '************ PROPOSED VALUE projeccpropvalue***************************************************************
            
            For EX = 1 To 5
                Debug.Print Trim("ENGY_ITEM_PROPOSED_0" & EX)
                If varAttributes(iattr_Insp).TagString = Trim("ENGY_ITEM_PROPOSED_0" & EX) Then
                    If IsNull(!projeccpropvalue) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes(iattr_Insp).textString = ""
                    Else
                        varAttributes(iattr_Insp).textString = SEPTEXT2(!projeccpropvalue, EX, 5, 24)
                    End If
                End If
            Next EX
            '************ CODE VALUE projcodevalue***************************************************************
            For EX = 1 To 5
                Debug.Print Trim("ENGY_ITEM_REQUIRED_0" & EX)
                If varAttributes(iattr_Insp).TagString = Trim("ENGY_ITEM_REQUIRED_0" & EX) Then
                    If IsNull(!projcodevalue) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes(iattr_Insp).textString = ""
                    Else
                        varAttributes(iattr_Insp).textString = SEPTEXT2(!projcodevalue, EX, 5, 54)
                    End If
                End If
            Next EX
            '************ SUPPORTING DOCUMENT projsuppdoc***************************************************************
            For EX = 1 To 5
                Debug.Print Trim("ENGY_SUP_DOC_0" & EX)
                If varAttributes(iattr_Insp).TagString = Trim("ENGY_SUP_DOC_0" & EX) Then
                    If IsNull(!projsuppdoc) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes(iattr_Insp).textString = ""
                    Else
                        varAttributes(iattr_Insp).textString = SEPTEXT2(!projsuppdoc, EX, 5, 18)
                    End If
                End If
            Next EX
            
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            '************ END TEXT INSERT LOOP ***************************************************************
    
        Next iattr_Insp
         TblPT(1) = TblPT(1) - nJumpRow
        .MoveNext
        
    Next l
    rstProjECCItem.Close
End With

Set DB = Nothing

Exit_writeENERGYTABULAR:
frmInsp.Hide
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub writeEnergyTabularXT()


'############################################################################
'############################################################################
'#         01/30/17
'#
'#
'#
'#
'#
'############################################################################
'############################################################################

Dim DB As Database
Dim rstProjECCItem As Recordset
Dim strProjECCItem As String
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

Dim nInspHeadDwg, nInspRowDwg, nInspRowDwgXT As Variant
Dim nJumpHead, nJumpRow As Integer
Dim blockRefObj, blockRefObj1 As AcadBlockReference
Dim BlkObj As Object
Dim k, l As Integer


CreateCNstyle
TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify

'nInspHeadDwg = "H:\drawfile\jvba\X_INSP_HEAD_10.dwg"
'nInspRowDwg = "H:\drawfile\jvba\X_INSP_ROW_10.dwg"
nInspHeadDwg = "\\jwja-svr-10\drawfile\jvba\XENGYHD05.dwg"
nInspRowDwg = "\\jwja-svr-10\drawfile\jvba\XENGYITEM05.dwg"
nInspRowDwgXT = "\\jwja-svr-10\drawfile\jvba\XENGYITEM05XXT.dwg"
    
    
nJumpHead = 96 '11.5 X 12 ; 8 X 12
'nJumpRow = 42 '3.5 X 12
nJumpRow = 48 '5 X 12 ; 3 X 12 ; 4 X 12
'InsertJBlock "H:\drawfile\jvba\X_WORKITEM_HEAD02.dwg", tblPT, "Model"

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_DOB_INSPECTIONS\X-dob_inspections_02.mdb")

xProjECCId = frmInsp.seleProjECC
Debug.Print xProjECCId
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

'strProjInsp = "SELECT PROJINSPITEM.proj_no, PROJINSPITEM.projinspid, PROJINSPITEM.INPUT, PROJINSPITEM.INSPNOTE, " _
& "INSPECTION.INSPTYPE, INSPECTION.CODENO, INSPECTION.INSPNAME, INSPECTION.LEADDAYS, PROJINSPITEM.INSTRUCTION, " _
& "INSPECTION.INSPCRITERIOR, INSPECTION.INSPFREQ, INSPECTION.INSPREF, INSPECTION.INSPCITATION, INSPECTION.ITEM_INSTRUCTION " _
& "FROM INSPECTION INNER JOIN PROJINSPITEM ON INSPECTION.INSPID = PROJINSPITEM.INSPID " _
& "WHERE (((PROJINSPITEM.projinspid) = " & xInspId & ") AND ((INSPECTION.INSPTYPE)= 'ENERGY'))" _
& "ORDER BY INSPECTION.INSPTYPE, INSPECTION.CODENO, INSPECTION.INSPID;"

strProjECCItem = "SELECT PROJECC.CLIMATEZONE, PROJECC.CODE, PROJECC.ENERGYWORKSCOPE, PROJECC.PROJ_NO, PROJECCITEM.projeccitemid, PROJECCITEM.projeccid, " _
& "PROJECCITEM.ecctabid, PROJECCITEM.projeccdesc, PROJECCITEM.projeccpropvalue, " _
& "PROJECCITEM.projcodevalue, PROJECCITEM.projsuppdoc, PROJECCITEM.projnote, " _
& "PROJECCITEM.input, PROJECCITEM.projcodeid, PROJECCITEM.projcitation, PROJECCITEM.projprovision " _
& "FROM PROJECC INNER JOIN PROJECCitem ON PROJECC.projeccid = PROJECCITEM.projeccid  " _
& "WHERE (((PROJECCITEM.projeccid)= " & xProjECCId & ")) " _
& "ORDER BY PROJECCITEM.projcitation;"


Set rstProjECCItem = DB.OpenRecordset(strProjECCItem)
rstProjECCItem.MoveFirst
rstProjECCItem.MoveLast
rstProjECCItem.MoveFirst
If rstProjECCItem.RecordCount = 0 Then
    GoTo Exit_writeENERGYTABULAR
End If
nInspRow = rstProjECCItem.RecordCount

With rstProjECCItem
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'InsertJBlock "Z:\jvba\XENGYHD04.dwg", tblPT, "Model"
    Set blockRefObj1 = ThisDrawing.ModelSpace.InsertBlock _
                 (TblPT, nInspHeadDwg, 1, 1, 1, 0) 'First dwg
    TblPT(1) = TblPT(1) - nJumpHead
    Dim varAttributes1 As Variant
    'Dim EX As Integer
    varAttributes1 = blockRefObj1.GetAttributes
    'Debug.Print !PROJECC.CODE
    For iattr_Insp = LBound(varAttributes1) To UBound(varAttributes1)
    '************ PROJ_NO ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "XPROJ_NO" Then
                If IsNull(!PROJ_NO) = True Then
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = !PROJ_NO
                    Debug.Print varAttributes1(iattr_Insp).textString
                End If
            End If
    '************ XDATE ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "XDATE" Then
                'If IsNull(!CODENO) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    'varAttributes(iattr_Insp).TextString = ""
                'Else
                    varAttributes1(iattr_Insp).textString = Date
                    
                'End If
            End If
    '************ CLIMATE ZONE ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "CLIMATEZONE" Then
                If IsNull(!climatezone) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = !climatezone
                    
                End If
            End If
    '************ CODE ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "CODE" Then
                If IsNull(!CODE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = !CODE
                    
                End If
            End If
    '************ SCOPE01 ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "SCOPE01" Then
                If IsNull(!ENERGYWORKSCOPE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = SEPTEXT2(!ENERGYWORKSCOPE, 1, 2, 40)
                    
                End If
            End If
    '************ SCOPE02 ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "SCOPE02" Then
                If IsNull(!ENERGYWORKSCOPE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = SEPTEXT2(!ENERGYWORKSCOPE, 2, 2, 40)
                    
                End If
            End If
    Next iattr_Insp
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    'MPROJinspID = !PROJinspID
    For l = 0 To nInspRow - 1
        Dim mLineCount, mChrPerLine As Integer
        Dim mTotalLine As Integer
        mTotalLine = 5
        mChrPerLine = 30 'PROPOSED VALUE projeccpropvalue WAS 24
        tmpChrNo = Len(!projeccpropvalue)
        'projcodevalue
        tmpChrNo2 = Len(!projcodevalue)
        tmpLineCount = tmpChrNo / mChrPerLine
        tmpLineCount2 = tmpChrNo2 / 54
        If tmpLineCount - tmpLineCount2 > 0 Then 'use 1 or 2
        tmpLineCount = tmpLineCount
        Else
        tmpLineCount = tmpLineCount2
        End If
        Debug.Print tmpLineCount
        If tmpLineCount > 6 Then
        Debug.Print tmpChrNo
        End If
        tmpLineCount = Round(tmpLineCount)
        If tmpLineCount <= 5 Then
            mTotalLine = 5
            mLineCount = 5
        End If
        If tmpLineCount > 5 Then
            mTotalLine = 10
            mLineCount = 10
        End If



        If mLineCount >= 6 Then
           
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, nInspRowDwgXT, 1, 1, 1, 0) 'First dwg
                     'tblPT(1) = tblPT(1) - nJumpRow
        
        Else
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                 (TblPT, nInspRowDwg, 1, 1, 1, 0) 'First dwg
                 
        End If
        Dim varAttributes As Variant
        Dim EX As Integer
        varAttributes = blockRefObj.GetAttributes
    
        For iattr_Insp = LBound(varAttributes) To UBound(varAttributes)
    
            Debug.Print varAttributes(iattr_Insp).TagString
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            '************ CITATATION ***************************************************************
            For EX = 1 To mLineCount
                Debug.Print Trim("ENGY_ITEM_CIT_0" & EX)
                If varAttributes(iattr_Insp).TagString = Trim("ENGY_ITEM_CIT_0" & EX) Then
                    If IsNull(!projcitation) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes(iattr_Insp).textString = ""
                    Else
                        Debug.Print !projcitation & "*********************************" & EX
                        varAttributes(iattr_Insp).textString = SEPTEXT2(!projcitation, EX, mTotalLine, 12)
                    End If
                End If
            Next EX
            '************ PROVISION projprovision***************************************************************
            For EX = 1 To mLineCount
                Debug.Print Trim("ENGY_ITEM_PROV_0" & EX)
                If varAttributes(iattr_Insp).TagString = Trim("ENGY_ITEM_PROV_0" & EX) Then
                    If IsNull(!projprovision) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes(iattr_Insp).textString = ""
                    Else
                        Debug.Print !projprovision & "*********************************" & EX
                        varAttributes(iattr_Insp).textString = SEPTEXT2(!projprovision, EX, mTotalLine, 22)
                    End If
                End If
            Next EX
            
            '************ DESCRIPTION projeccdesc***************************************************************
            tmpChrNo = Len(!projeccdesc)
            
            For EX = 1 To 5
                Debug.Print Trim("ENGY_ITEM_DESC_0" & EX)
                If varAttributes(iattr_Insp).TagString = Trim("ENGY_ITEM_DESC_0" & EX) Then
                    If IsNull(!projeccdesc) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes(iattr_Insp).textString = ""
                    Else
                        Debug.Print !projeccdesc & "*********************************" & EX
                        varAttributes(iattr_Insp).textString = SEPTEXT2(!projeccdesc, EX, mTotalLine, 28)
                    End If
                End If
            Next EX
            '************ PROPOSED VALUE projeccpropvalue***************************************************************
            For EX = 1 To mLineCount
                Debug.Print Trim("ENGY_ITEM_PROPOSED_0" & EX)
                If varAttributes(iattr_Insp).TagString = Trim("ENGY_ITEM_PROPOSED_0" & EX) Then
                    If IsNull(!projeccpropvalue) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes(iattr_Insp).textString = ""
                    Else
                        Debug.Print !projeccpropvalue & "*********************************" & EX
                        varAttributes(iattr_Insp).textString = SEPTEXT2(!projeccpropvalue, EX, mTotalLine, mChrPerLine)
                        If EX = 6 Then
                        TblPT(1) = TblPT(1) - nJumpRow
                        End If
                    End If
                End If
            Next EX
            '************ CODE VALUE projcodevalue***************************************************************
            For EX = 1 To mLineCount
                Debug.Print Trim("ENGY_ITEM_REQUIRED_0" & EX)
                If varAttributes(iattr_Insp).TagString = Trim("ENGY_ITEM_REQUIRED_0" & EX) Then
                    If IsNull(!projcodevalue) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes(iattr_Insp).textString = ""
                    Else
                        varAttributes(iattr_Insp).textString = SEPTEXT2(!projcodevalue, EX, mTotalLine, 54)
                    End If
                End If
            Next EX
            '************ SUPPORTING DOCUMENT projsuppdoc***************************************************************
            For EX = 1 To mLineCount
                Debug.Print Trim("ENGY_SUP_DOC_0" & EX)
                If varAttributes(iattr_Insp).TagString = Trim("ENGY_SUP_DOC_0" & EX) Then
                    If IsNull(!projsuppdoc) = True Then '++++++++++++++++++++++++++++++++++++++++++
                        varAttributes(iattr_Insp).textString = ""
                    Else
                        varAttributes(iattr_Insp).textString = SEPTEXT2(!projsuppdoc, EX, mTotalLine, 18)
                    End If
                End If
            Next EX
            
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            '************ END TEXT INSERT LOOP ***************************************************************
    
        Next iattr_Insp
         TblPT(1) = TblPT(1) - nJumpRow
        .MoveNext
        
    Next l
    rstProjECCItem.Close
End With

Set DB = Nothing

Exit_writeENERGYTABULAR:
frmInsp.Hide
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub writeEnergyTabularNW()


'############################################################################
'############################################################################
'#         01/30/17 10/5/18
'#
'#
'#
'#
'#
'############################################################################
'############################################################################

Dim DB As Database
Dim rstProjECCItem As Recordset
Dim strProjECCItem, tabText, mStyleName As String
Dim nInspRow As Integer    'ProjSet no.
Dim iattr_Insp As Integer
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
Dim hLineStart(0 To 2) As Double
Dim TextPosition(0 To 2) As Double
Dim rowHt(0 To 1) As Double
Dim vBar(6) As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
Dim nInspHeadDwg, nInspRowDwg, nInspRowDwgXT As Variant
Dim nJumpHead, nJumpRow As Integer
Dim blockRefObj, blockRefObj1 As AcadBlockReference
Dim BlkObj As Object
Dim k, i As Integer
Dim wCitation, wProvision, wDesc, wDesignValue, wCodeValue, wSupDoc As Double

'Dim tmpArrayAss(0 To 1) As Double
'Dim rcArray As Variant
'Dim tblWidth As Double
'Dim height As Double
'Dim mode As Long
'Dim prompt As String
If frmInsp.OptionButton6.value Then
CreateCNstyle
mStyleName = "cn"
txtHt = 4
ElseIf frmInsp.OptionButton7.value Then
CreateARstyle
mStyleName = "ar"
txtHt = 6
End If
If frmInsp.OptionButton1.value Then
    mSCF = 48 'mSCF= Scale Factor
    If frmInsp.OptionButton6.value Then
     
     xScF = mSCF / 12

    ElseIf frmInsp.OptionButton7.value Then
        
        xScF = (mSCF / 12) * 1.5
    End If

    
    'width = mSCF * 8 '48 x 8 = 384
    wCitation = 49 ' 4'-1"
    wProvision = 80 '6'-8",
    wDesc = 80
    wDesignValue = 88
    wCodeValue = 144
    wSupDoc = 63
    

ElseIf frmInsp.OptionButton2.value Then
    mSCF = 96
    wCitation = 49 * 2 ' 4'-1"
    wProvision = 80 * 2 '6'-8",
    wDesc = 80 * 2
    wDesignValue = 88 * 2
    wCodeValue = 144 * 2
    wSupDoc = 63 * 2
    xScF = mSCF / 12
    xScFBig = xScF * 1.5
ElseIf frmInsp.OptionButton3.value Then
    mSCF = 64
    'width = mSCF * 7 '64 x 8 = 512
    xScF = mSCF / 12
ElseIf frmInsp.OptionButton4.value Then
    mSCF = 128
    'width = mSCF * 7 '64 x 8 = 512
    xScF = mSCF / 12
ElseIf frmInsp.OptionButton5.value Then
    mSCF = 192
    'width = mSCF * 7 '64 x 8 = 512
    xScF = mSCF / 12

End If
xScFBig = xScF * 1.5
'*********************************************************************************************
'*********************************************************************************************
'startPoint(0) = 0#: startPoint(1) = 0#: startPoint(2) = 0#
'SectionPoint(0) = 0#: SectionPoint(1) = 0#: SectionPoint(2) = 0#
'paraPoint(0) = 0#: paraPoint(1) = 0#: paraPoint(2) = 0#
TextPosition(0) = 0#: TextPosition(1) = 0#: TextPosition(2) = 0#
'width = 384 '32' X 12'

CodeText = ""
mChapterNo = 1
'**********************************************************************************************************
'**********************************************************************************************************

TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0

txtRow = ""

mode = acAttributeModeVerify

'nInspHeadDwg = "H:\drawfile\jvba\X_INSP_HEAD_10.dwg"
'nInspRowDwg = "H:\drawfile\jvba\X_INSP_ROW_10.dwg"

nInspHeadDwg = "\\jwja-svr-10\drawfile\jvba\XENGYHD05.dwg"
nInspRowDwg = "\\jwja-svr-10\drawfile\jvba\XENGYITEM05.dwg"
nInspRowDwgXT = "\\jwja-svr-10\drawfile\jvba\XENGYITEM05XXT.dwg"
nENGYFOOTER = "\\jwja-svr-10\drawfile\jvba\XENGYFT03.dwg"
nJumpHead = 96 '11.5 X 12 ; 8 X 12
'nJumpRow = 42 '3.5 X 12
nJumpRow = 48 '5 X 12 ; 3 X 12 ; 4 X 12
'InsertJBlock "H:\drawfile\jvba\X_WORKITEM_HEAD02.dwg", tblPT, "Model"
indLeft = 6
indright = 4
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_DOB_INSPECTIONS\X-dob_inspections_02.mdb")

xProjECCId = frmInsp.seleProjECC
Debug.Print xProjECCId

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

'strProjInsp = "SELECT PROJINSPITEM.proj_no, PROJINSPITEM.projinspid, PROJINSPITEM.INPUT, PROJINSPITEM.INSPNOTE, " _
& "INSPECTION.INSPTYPE, INSPECTION.CODENO, INSPECTION.INSPNAME, INSPECTION.LEADDAYS, PROJINSPITEM.INSTRUCTION, " _
& "INSPECTION.INSPCRITERIOR, INSPECTION.INSPFREQ, INSPECTION.INSPREF, INSPECTION.INSPCITATION, INSPECTION.ITEM_INSTRUCTION " _
& "FROM INSPECTION INNER JOIN PROJINSPITEM ON INSPECTION.INSPID = PROJINSPITEM.INSPID " _
& "WHERE (((PROJINSPITEM.projinspid) = " & xInspId & ") AND ((INSPECTION.INSPTYPE)= 'ENERGY'))" _
& "ORDER BY INSPECTION.INSPTYPE, INSPECTION.CODENO, INSPECTION.INSPID;"

strProjECCItem = "SELECT PROJECC.CLIMATEZONE, PROJECC.CODE, PROJECC.ENERGYWORKSCOPE, PROJECC.PROJ_NO, PROJECCITEM.projeccitemid, PROJECCITEM.projeccid, " _
& "PROJECCITEM.ecctabid, PROJECCITEM.projeccdesc, PROJECCITEM.projeccpropvalue, " _
& "PROJECCITEM.projcodevalue, PROJECCITEM.projsuppdoc, PROJECCITEM.projnote, " _
& "PROJECCITEM.input, PROJECCITEM.projcodeid, PROJECCITEM.projcitation, PROJECCITEM.projprovision " _
& "FROM PROJECC INNER JOIN PROJECCitem ON PROJECC.projeccid = PROJECCITEM.projeccid  " _
& "WHERE (((PROJECCITEM.projeccid)= " & xProjECCId & ")) " _
& "ORDER BY PROJECCITEM.projcitation;"



Set rstProjECCItem = DB.OpenRecordset(strProjECCItem)
rstProjECCItem.MoveFirst
rstProjECCItem.MoveLast
rstProjECCItem.MoveFirst
If rstProjECCItem.RecordCount = 0 Then
    GoTo Exit_writeENERGYTABULAR
End If
nInspRow = rstProjECCItem.RecordCount

With rstProjECCItem
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'InsertJBlock "Z:\jvba\XENGYHD04.dwg", tblPT, "Model"
    '*******************************************************************************************************
    '**************************   Write Head Block *********************************************************
    '*******************************************************************************************************
    Set blockRefObj1 = ThisDrawing.ModelSpace.InsertBlock _
    (TblPT, nInspHeadDwg, 1, 1, 1, 0) 'First dwg
    
    Dim varAttributes1 As Variant
    'Dim EX As Integer
    varAttributes1 = blockRefObj1.GetAttributes
    'Debug.Print !PROJECC.CODE
    For iattr_Insp = LBound(varAttributes1) To UBound(varAttributes1)
    '************ PROJ_NO ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "XPROJ_NO" Then
                If IsNull(!PROJ_NO) = True Then
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = !PROJ_NO
                    Debug.Print varAttributes1(iattr_Insp).textString
                End If
            End If
    '************ XDATE ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "XDATE" Then
                'If IsNull(!CODENO) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    'varAttributes(iattr_Insp).TextString = ""
                'Else
                    varAttributes1(iattr_Insp).textString = Date
                    
                'End If
            End If
    '************ CLIMATE ZONE ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "CLIMATEZONE" Then
                If IsNull(!climatezone) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = !climatezone
                    
                End If
            End If
    '************ CODE ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "CODE" Then
                If IsNull(!CODE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = !CODE
                    
                End If
            End If
    '************ SCOPE01 ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "SCOPE01" Then
                If IsNull(!ENERGYWORKSCOPE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = SEPTEXT2(!ENERGYWORKSCOPE, 1, 2, 12)
                    
                End If
            End If
    '************ SCOPE02 ***************************************************************
            If varAttributes1(iattr_Insp).TagString = "SCOPE02" Then
                If IsNull(!ENERGYWORKSCOPE) = True Then '++++++++++++++++++++++++++++++++++++++++++
                    varAttributes1(iattr_Insp).textString = ""
                Else
                    varAttributes1(iattr_Insp).textString = SEPTEXT2(!ENERGYWORKSCOPE, 2, 2, 12)
                    
                End If
            End If
    '************ SCOPE02 ***************************************************************
    Next iattr_Insp
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    TblPT(1) = TblPT(1) - nJumpHead 'TblPT(1)=0
    mTab = 12
    TextPosition(0) = TblPT(0)
    TextPosition(1) = TblPT(1)
    oldrowht = TblPT(1)
    For i = 0 To nInspRow - 1 'ECC Item Count
    oldrowht = rowHt(1)
    '+++++++++++++++++++++++++++++++
    '+++++++ 1 ProjCitation+++++++++++
    '+++++++++++++++++++++++++++++++
    tabText = ""
    
    tabText = "\P" & rstProjECCItem!projcitation & "\P"
    TextPosition(0) = TextPosition(0) + indLeft
    'TextPosition(1) = TextPosition(1)
    TextPosition(1) = (Int(TextPosition(1) / 4) * 4) - 4
    vBar(0) = TextPosition(0) - indLeft
    Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, wCitation - indLeft - indright, tabText)
    
    mtextObj.StyleName = mStyleName
    mtextObj.height = xScF
    
    mtextObj.Update
    mtextObj.GetBoundingBox minText, maxText
    'TextPosition(1) = minText(1)
    rowHt(0) = minText(0): rowHt(1) = minText(1) 'Set arry for minText
    Debug.Print "1) " & minText(0) & ":" & minText(1)
    'If TextPosition(1) > minText(1) Then
    rowHt(1) = minText(1)
    Debug.Print "1st. Ht.: " & rowHt(1)
    'End If

    '+++++++++++++++++++++++++++++++
    '+++++++ 2 projprovision+++++++++++
    '+++++++++++++++++++++++++++++++
    tabText = ""
    tabText = "\P" & rstProjECCItem!projprovision & "\P"
    TextPosition(0) = wCitation + indLeft
    vBar(1) = TextPosition(0) - indLeft
    'TextPosition(1) = TextPosition(1)
    'TextPosition(1) = (Int(TextPosition(1) / 4) * 4) - 4
    Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, wProvision - indLeft - indright, tabText)
    mtextObj.StyleName = mStyleName
    mtextObj.height = xScF
    mtextObj.Update
    mtextObj.GetBoundingBox minText, maxText
    'TextPosition(1) = minText(1)
    'rowHt(0) = minText(0): rowHt(1) = minText(1) 'Set arry for minText
    Debug.Print "2) " & minText(0) & ":" & minText(1)
    If rowHt(1) > minText(1) Then
    Debug.Print "minText: " & minText(1)
    rowHt(1) = minText(1)
    Debug.Print "Evaled: " & rowHt(1)
    End If
    '+++++++++++++++++++++++++++++++
    '+++++++ 3 projeccdesc+++++++++++
    '+++++++++++++++++++++++++++++++
    tabText = ""
    tabText = "\P" & rstProjECCItem!projeccdesc & "\P"
    TextPosition(0) = wCitation + wProvision + indLeft
    vBar(2) = TextPosition(0) - indLeft
    'TextPosition(1) = TextPosition(1)
    'TextPosition(1) = (Int(TextPosition(1) / 4) * 4) - 4
    Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, wDesc - indLeft - indright, tabText)
    mtextObj.StyleName = mStyleName
    mtextObj.height = xScF
    mtextObj.Update
    mtextObj.GetBoundingBox minText, maxText
    'TextPosition(1) = minText(1)
    'rowHt(0) = minText(0): rowHt(1) = minText(1)  'Set arry for minText
    Debug.Print "3) " & minText(0) & ":" & minText(1)
    If rowHt(1) > minText(1) Then
    Debug.Print "minText: " & minText(1)
    rowHt(1) = minText(1)
    Debug.Print "Evaled: " & rowHt(1)
    End If
    '+++++++++++++++++++++++++++++++
    '+++++++ 4 projeccpropvalue+++++++++++
    '+++++++++++++++++++++++++++++++
    tabText = ""
    tabText = "\P" & rstProjECCItem!projeccpropvalue & "\P"
    TextPosition(0) = wCitation + wProvision + wDesc + indLeft
    vBar(3) = TextPosition(0) - indLeft
    'TextPosition(1) = TextPosition(1)
    'TextPosition(1) = (Int(TextPosition(1) / 4) * 4) - 4
    Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, wDesignValue - indLeft - indright, tabText)
    mtextObj.StyleName = mStyleName
    mtextObj.height = xScF
    mtextObj.Update
    mtextObj.GetBoundingBox minText, maxText
    'TextPosition(1) = minText(1)
    'rowHt(0) = minText(0): rowHt(1) = minText(1)  'Set arry for minText
    Debug.Print "4) " & minText(0) & ":" & minText(1)
    If rowHt(1) > minText(1) Then
    Debug.Print "minText: " & minText(1)
    rowHt(1) = minText(1)
    Debug.Print "Evaled: " & rowHt(1)
    End If
    '+++++++++++++++++++++++++++++++
    '+++++++ 5 projcodevalue+++++++++++
    '+++++++++++++++++++++++++++++++
    tabText = ""
    tabText = "\P" & rstProjECCItem!projcodevalue & "\P"
    TextPosition(0) = wCitation + wProvision + wDesc + wDesignValue + indLeft
    vBar(4) = TextPosition(0) - indLeft
    'TextPosition(1) = TextPosition(1)
    'TextPosition(1) = (Int(TextPosition(1) / 4) * 4) - 4
    Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, wCodeValue - indLeft - indright, tabText)
    mtextObj.StyleName = mStyleName
    mtextObj.height = xScF
    mtextObj.Update
    mtextObj.GetBoundingBox minText, maxText
    'TextPosition(1) = minText(1)
    'rowHt(0) = minText(0): rowHt(1) = minText(1) 'Set arry for minText
    Debug.Print "5) " & minText(0) & ":" & minText(1)
    If rowHt(1) > minText(1) Then
    Debug.Print "minText: " & minText(1)
    rowHt(1) = minText(1)
    Debug.Print "Evaled: " & rowHt(1)
    End If
    '+++++++++++++++++++++++++++++++
    '+++++++ 6 projsuppdoc+++++++++++
    '+++++++++++++++++++++++++++++++
    tabText = ""
    tabText = "\P" & rstProjECCItem!projsuppdoc & "\P"
    TextPosition(0) = wCitation + wProvision + wDesc + wDesignValue + wCodeValue + indLeft
    vBar(5) = TextPosition(0) - indLeft
    vBar(6) = TextPosition(0) + wSupDoc - indLeft
    'TextPosition(1) = TextPosition(1)
    'TextPosition(1) = (Int(TextPosition(1) / 4) * 4) - 4
    Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, wSupDoc - indLeft - indright, tabText)
    mtextObj.StyleName = mStyleName
    mtextObj.height = xScF
    mtextObj.Update
    mtextObj.GetBoundingBox minText, maxText
    'TextPosition(1) = minText(1)
    'rowHt(0) = minText(0): rowHt(1) = minText(1) 'Set arry for minText
    Debug.Print "6) " & minText(0) & ":" & minText(1)
    If rowHt(1) > minText(1) Then
    Debug.Print "minText: " & minText(1)
    rowHt(1) = minText(1)
    Debug.Print "Evaled: " & rowHt(1)
    End If
    '+++++++++++++++++++++++++++++++
    '+++++++++++++++++++++++++++++++
    '+++++++++++++++++++++++++++++++
    
    TextPosition(0) = 0#: TextPosition(1) = rowHt(1)
    Debug.Print "TextPosition(1) AFter loop: " & TextPosition(1)
    '+++++++++++++++++++++++++++++++++++++++
    '+++++++++++++++++++++++++++++++++++++++
    '+++++++++++++++++++++++++++++++++++++++
    TblPT(0) = TblPT(0)
    If i = 0 Then
        TblPT(1) = -nJumpHead
    Else
        'TblPT(1) = oldrowht
        TblPT(1) = (Int(oldrowht / 4) * 4) - 4
    End If
    tblPT2(1) = rowHt(1)
    '_________________________
    tmpY = Int(tblPT2(1) / 4)
    tblPT2(1) = (tmpY * 4) - 4
     '_________________________
    
    'TblPT(1) = rowHt(1) + oldrowht
    'oldrowht = TextPosition(1)
        For k = LBound(vBar) To UBound(vBar)
        
         
         'TblPT(0) = vBar(k): TblPT(1) = rowHt(1) + oldrowht
         TblPT(0) = vBar(k)
         tblPT2(0) = vBar(k)
         
         ThisDrawing.ModelSpace.AddLine TblPT, tblPT2
         
        
    
        Next k
    '+++++++++++++++++++++++++++++++++++++++
    '+++++++++++++++++++++++++++++++++++++++
    '+++++++++++++++++++++++++++++++++++++++
    'TextPosition(0) = 0#: TextPosition(1) = 0#: TextPosition(2) = 0#
    hLineStart(0) = 0#: hLineStart(1) = tblPT2(1): hLineStart(2) = 0#
    ThisDrawing.ModelSpace.AddLine hLineStart, tblPT2
    
    .MoveNext
    
        
    Next i
'===============================================================================
' FOOTER ENERGY
'===============================================================================

     Set blockRefObj1 = ThisDrawing.ModelSpace.InsertBlock _
    (hLineStart, nENGYFOOTER, 1, 1, 1, 0) 'First dwg
'===============================================================================

End With
Exit_writeENERGYTABULAR:
rstProjECCItem.Close
Set DB = Nothing

Exit_writeEnergyTabularNW:
frmInsp.Hide
End Sub
Public Function MaxValOfIntArray(ByRef TheArray As Variant) As Double
'This function gives max value of int array without sorting an array
Dim i As Integer
Dim MaxIntegersIndex As Integer
MaxIntegersIndex = 0

For i = 1 To UBound(TheArray)
    If TheArray(i) > TheArray(MaxIntegersIndex) Then
        MaxIntegersIndex = i
    End If
Next
'index of max value is MaxValOfIntArray
MaxValOfIntArray = TheArray(MaxIntegersIndex)
End Function

'Function Lowest(ByRef TheArray As Double) As Double
        'Dim Result As Double = Double.MaxValue
        'For Each D As Double In Array
            'If D <= Result Then Result = D
        'Next
        'Return Result
    'End Function
