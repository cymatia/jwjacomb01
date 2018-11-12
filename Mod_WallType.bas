Attribute VB_Name = "Mod_WallType"
Public Sub ShowProjWallType()
frmWallType.show
End Sub
Function writeProjENV(mProjEnvID As Integer)
'On Error GoTo Err_writemtext
Dim DB As Database
Dim rstProjBldgEnv, rstProjEnv, rstEnv, rstMAat, rstProjENVUA, rs As Recordset
Dim strProjBldgEnv, strProjEnv, strEnv, strMat, strProjENVUA As String
Dim rnProjBldgEnv, rnProjEnv, rnEnv, rnMat As Integer
Dim iProjBldgEnv, iProjEnv, iEnv, iMat  As Integer
Dim width As Double
Dim tmpEnvCategory As String
Dim MatArray() As Variant
'Dim StartPoint(0 To 2) As Double
'Dim SectionPoint(0 To 2) As Double
'Dim paraPoint(0 To 2) As Double
'Dim TextPosition(0 To 2) As Double
Dim TblPT(0 To 2) As Double
'Dim TblEndPT(0 To 2) As Double
'Dim SizestartPoint(0 To 2) As Double
'Dim width As Double
'Dim mATno As Integer

Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute


Dim nInspHeadDwg, nInspRowDwg As Variant
Dim nJumpHead, nJumpRow As Integer
Dim blockRefObj As AcadBlockReference
Dim BlkObj As Object
Dim k, l As Integer

'*********************************************************************************************
'*********************************************************************************************
Set DB = OpenDatabase("\\jwja-svr-10\jdb\DB_WALL\WALLASS03.mdb")
'*********************************************************************************************
'*********************************************************************************************
' 09/22/18
'*********************************************************************************************
'*********************************************************************************************
'mSCF= Scale Factor
If frmWallType.OptionButton1.value Then
    mSCF = 48 'mSCF= Scale Factor
    'width = mSCF * 8 '48 x 8 = 384
    width = mSCF * 7 '48 x 7 = 336
    xScF = mSCF / 12
ElseIf frmProjCode.OptionButton2.value Then
    mSCF = 96
    width = mSCF * 7 '96 x 8 = 768
    xScF = mSCF / 12
    xScFBig = xScF * 1.5
ElseIf frmProjCode.OptionButton3.value Then
    mSCF = 64
    width = mSCF * 7 '64 x 8 = 512
    xScF = mSCF / 12
ElseIf frmProjCode.OptionButton4.value Then
    mSCF = 128
    width = mSCF * 7 '64 x 8 = 512
    xScF = mSCF / 12
ElseIf frmProjCode.OptionButton5.value Then
    mSCF = 192
    width = mSCF * 7 '64 x 8 = 512
    xScF = mSCF / 12

End If
xScFBig = xScF * 1.5


    'mSCF = 48 'mSCF= Scale Factor
    'width = mSCF * 8 '48 x 8 = 384
    'width = 1440 * 6

'*********************************************************************************************
'*********************************************************************************************
'startPoint(0) = 0#: startPoint(1) = 0#: startPoint(2) = 0#
'SectionPoint(0) = 0#: SectionPoint(1) = 0#: SectionPoint(2) = 0#
'paraPoint(0) = 0#: paraPoint(1) = 0#: paraPoint(2) = 0#
'TextPosition(0) = 0#: TextPosition(1) = 0#: TextPosition(2) = 0#
'width = 384 '32' X 12'
TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0

'codetext = ""
'mChapterNo = 1
'**********************************************************************************************************
'**********************************************************************************************************
mode = acAttributeModeVerify
nWTPROJENVHD = "\\jwja-svr-10\drawfile\jvba\XWTPROJENVHD.dwg"
nWTENVSUMHD = "\\jwja-svr-10\drawfile\jvba\XWTENVSUMHD.dwg"
nWTMATROW01 = "\\jwja-svr-10\drawfile\jvba\XWTMATROW01.dwg"
nWTENVSUMFOOT = "\\jwja-svr-10\drawfile\jvba\XWTENVSUMFOOT.dwg"
nWTPROJENVUAHD = "\\jwja-svr-10\drawfile\jvba\XWTPROJENVUAHD.dwg"
nWTPROJENVUArow = "\\jwja-svr-10\drawfile\jvba\XWTPROJENVUAROW.dwg"
nWTPROJENVUAFOOT = "\\jwja-svr-10\drawfile\jvba\XWTPROJENVUAFOOT.dwg"

nJumpPROJENVHD = 30 '2.5 X 12
nJumpENVsumHD = 24 '2 X 12
nJumpMATROW01 = 18 '1.5 X 12
nJumpEnvSumFoot = 30 '2.5 X 12
nJumpProjEnvUAHD = 42 '3.5 X 12
nJumpProjEnvUArow = 18 '1.5 X 12
nJumpPRJENVUAFOOT = 48 '4 X 12
nShift = 189 '15'-9x12
xProjEnvID = mProjEnvID
'CAD3
strProjEnvItem = "SELECT ConnProjBldgEnv.ProjBldgEnvid, ConnProjBldgEnv.ProjENvID, ProjEnvItem.ProjEnvNo, " _
    & "ProjEnvItem.ProjENvCategory, ProjEnvItem.ProjEnvType, ProjEnvItem.ProjENvKey,ProjEnvItem.ProjENvName,ProjEnvItem.ProjEnvArea, ProjEnvItem.ProjENvR, " _
    & "ProjEnvItem.ProjEnvU, ProjEnvItem.ProjUWt, ProjEnvItem.ProjRwt, ProjEnvItem.ProjEnvKey, ProjEnvItem.ProjComponentType, " _
    & "ProjEnvItem.ProjEnvUAdiff, ProjEnvItem.Utable, ProjEnvItem.ProjEnvUAtable, ProjEnvItem.ProjEnvFLdiff, ProjEnvItem.ProjEnvF, " _
    & "ProjEnvItem.ProjEnvFL, ProjEnvItem.ProjEnvFLtable, ProjEnvItem.Ftable, ProjEnvItem.ProjEnvCAdiff, ProjEnvItem.ProjEnvCA, " _
    & "ProjEnvItem.ProjEnvThick, ProjEnvItem.ProjEnvThickFrac, ProjEnvItem.ProjEnvNote, ProjEnvItem.ThermBrid, " _
    & "ProjEnvItem.ProjEnvAreaUnit, ProjEnvItem.ProjEnvUATotal, ProjEnvItem.CitationUtable " _
    & "FROM ProjEnvItem INNER JOIN ConnProjBldgEnv ON ProjEnvItem.ProjEnvID = ConnProjBldgEnv.ProjENvID " _
    & "WHERE (((ConnProjBldgEnv.ProjBldgEnvid)=" & xProjEnvID & ")) " _
    & "ORDER BY ProjEnvItem.ProjENvCategory, ProjEnvItem.ProjEnvType, ProjEnvItem.ProjEnvNo;"
'CAD5


        'strProjEnvItem = "SELECT ConnEnv.ConnEnvID, ConnEnv.EnvItemID, ConnEnv.ProjEnvID, ProjEnvItem.Proj_No, ProjEnvItem.ProjEnvNo, " _
        & "ProjEnvItem.ProjENvCategory, ProjEnvItem.ProjEnvType, ProjEnvItem.ProjENvName, ProjEnvItem.ProjEnvArea, ProjEnvItem.ProjENvR, " _
        & "ProjEnvItem.ProjEnvU, ProjEnvItem.ProjUWt, ProjEnvItem.ProjRwt, ProjEnvItem.ProjEnvKey, ProjEnvItem.ProjComponentType, " _
        & "ProjEnvItem.ProjEnvUAdiff, ProjEnvItem.Utable, ProjEnvItem.ProjEnvUAtable, ProjEnvItem.ProjEnvFLdiff, ProjEnvItem.ProjEnvF, " _
        & "ProjEnvItem.ProjEnvFL, ProjEnvItem.ProjEnvFLtable, ProjEnvItem.Ftable, ProjEnvItem.ProjEnvCAdiff, ProjEnvItem.ProjEnvCA, " _
        & "ProjEnvItem.ProjEnvThick, ProjEnvItem.ProjEnvThickFrac, ProjEnvItem.ProjEnvNote, ProjEnvItem.ThermBrid, " _
        & "ProjEnvItem.ProjEnvAreaUnit, ProjEnvItem.ProjEnvUATotal, ProjEnvItem.CitationUtable " _
        & "FROM ProjEnvItem INNER JOIN ConnEnv ON ProjEnvItem.ProjEnvID=ConnEnv.ProjEnvID  " _
        & "WHERE (ConnEnv.ProjEnvID)=" & xProjEnvID

Set rstProjEnvItem = DB.OpenRecordset(strProjEnvItem)
Debug.Print rstProjEnvItem.RecordCount
'rstProjEnvItem.MoveFirst
rstProjEnvItem.MoveLast
rstProjEnvItem.MoveFirst
If rstProjEnvItem.RecordCount = 0 Then
    GoTo Exit_writeProjENV
End If
nProjEnvItemRow = rstProjEnvItem.RecordCount
Debug.Print "nProjEnvItemRow:???? " & nProjEnvItemRow
With rstProjEnvItem ' With S 1
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
' Loop 1 Start

    'For k = 0 To nProjEnvItemRow - 1 'Loop 1
    Do While Not rstProjEnvItem.EOF 'Loop 1
        '==============================================
        mProjEnvID = !ProjEnvID
        'mEnvItemID = !EnvItemID
        '==============================================
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++ 1 dwg +++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                 (TblPT, nWTPROJENVHD, 1, 1, 1, 0) 'First dwg
        TblPT(1) = TblPT(1) - nJumpPROJENVHD
        Dim varAttributes As Variant
        varAttributes = blockRefObj.GetAttributes
        '************ XPROJENVHD ***************************************************************
        
        
        For iattr_ProjEnvItem = LBound(varAttributes) To UBound(varAttributes)
            Debug.Print varAttributes(iattr_ProjEnvItem).TagString
            If varAttributes(iattr_ProjEnvItem).TagString = "XPROJENVHD" Then
                If IsNull(!ProjENvKey) = True Then
                    varAttributes(iattr_ProjEnvItem).textString = ""
                Else
                    varAttributes(iattr_ProjEnvItem).textString = !ProjENvKey
                    Debug.Print varAttributes(iattr_ProjEnvItem).textString
                End If
            End If
        '************ XPROJENVCATEGORY ***************************************************************
        'Dim rs As Recordset

    'sSQL = "SELECT [SC Date] FROM [Stock Conversion] WHERE SCID = " _
     & txtSCNumber.value
    sSQL = "SELECT [EnvCategory] FROM [EnvCategory] WHERE EnvCatID = " & !ProjENvCategory
    ' & txtSCNumber.value
     
    Set rs = DB.OpenRecordset(sSQL)
    tmpEnvCategory = rs![EnvCategory]

        'tmpEnvCategory = dlookup("EnvCategory", "EnvCategory", "[EnvCatID]=" & ProjENvCategory)
            If varAttributes(iattr_ProjEnvItem).TagString = "XENVCATEGORY" Then
                If IsNull(tmpEnvCategory) = True Then
                    varAttributes(iattr_ProjEnvItem).textString = ""
                Else
                    varAttributes(iattr_ProjEnvItem).textString = tmpEnvCategory
                    Debug.Print varAttributes(iattr_ProjEnvItem).textString
                End If
            End If
            rs.Close
        Set rs = Nothing
        
        '*********************************************************************************************
        Next iattr_ProjEnvItem
        
        
        '************ XENVHD ***************************************************************
        'mEnvItemID = !EnvItemID
        'strEnv = "SELECT EnvItem.EnvItemID, EnvItem.EnvItemKey, EnvItem.EnvItemName, EnvItem.EnvItemArea, EnvItem.EnvItemU, " _
        & "EnvItem.EnvItemNote, EnvItem.EnvCategory, EnvItem.EnvItemThick, EnvItem.EnvItemR, EnvItem.EnvItemThickFrac, EnvItem.ThermBrid, " _
        & "EnvItem.EnvItemLink, EnvItem.EnvTypeID " _
        & "FROM EnvItem " _
        & "WHERE (EnvItem.EnvItemID)=" & mEnvItemID
        
        'strEnv = "SELECT ConnEnv.ProjEnvID, ConnEnv.EnvItemID, EnvItem.EnvItemKey, EnvItem.EnvItemName, EnvItem.EnvItemArea, " _
        & "EnvItem.EnvItemU, EnvItem.EnvItemNote, EnvItem.EnvCategory, EnvItem.EnvItemThick, EnvItem.EnvItemR, EnvItem.EnvItemThickFrac, " _
        & "EnvItem.ThermBrid, EnvItem.EnvItemLink, EnvItem.EnvTypeID " _
        & "FROM EnvItem INNER JOIN ConnEnv ON EnvItem.EnvItemID = ConnEnv.EnvItemID " _
        & "WHERE (ConnEnv.ProjEnvID)=" & mProjEnvID
    
        'CAD5
        'strEnv = "SELECT ConnEnv.ConnEnvID, ConnEnv.EnvItemID, ConnEnv.ProjEnvID, ProjEnvItem.Proj_No, ProjEnvItem.ProjEnvNo, " _
        & "ProjEnvItem.ProjENvCategory, ProjEnvItem.ProjEnvType, ProjEnvItem.ProjENvName, ProjEnvItem.ProjEnvArea, ProjEnvItem.ProjENvR, " _
        & "ProjEnvItem.ProjEnvU, ProjEnvItem.ProjUWt, ProjEnvItem.ProjRwt, ProjEnvItem.ProjEnvKey, ProjEnvItem.ProjComponentType, " _
        & "ProjEnvItem.ProjEnvUAdiff, ProjEnvItem.Utable, ProjEnvItem.ProjEnvUAtable, ProjEnvItem.ProjEnvFLdiff, ProjEnvItem.ProjEnvF, " _
        & "ProjEnvItem.ProjEnvFL, ProjEnvItem.ProjEnvFLtable, ProjEnvItem.Ftable, ProjEnvItem.ProjEnvCAdiff, ProjEnvItem.ProjEnvCA, " _
        & "ProjEnvItem.ProjEnvThick, ProjEnvItem.ProjEnvThickFrac, ProjEnvItem.ProjEnvNote, ProjEnvItem.ThermBrid, " _
        & "ProjEnvItem.ProjEnvAreaUnit, ProjEnvItem.ProjEnvUATotal, ProjEnvItem.CitationUtable " _
        & "FROM ProjEnvItem INNER JOIN ConnEnv ON ProjEnvItem.ProjEnvID=ConnEnv.ProjEnvID  " _
        & "WHERE (ConnEnv.ProjEnvID)=" & mProjEnvID
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
       strEnv = "SELECT ConnEnv.ProjEnvID, ConnEnv.EnvItemID, EnvItem.EnvItemKey, EnvItem.EnvItemName, EnvItem.EnvItemArea, EnvItem.EnvItemU, " _
       & "EnvItem.EnvItemNote, EnvItem.EnvCategory, EnvItem.EnvItemThick, EnvItem.EnvItemR, EnvItem.EnvItemThickFrac, EnvItem.ThermBrid, " _
       & "EnvItem.EnvItemLink, EnvItem.EnvTypeID " _
        & "FROM EnvItem INNER JOIN ConnEnv ON EnvItem.EnvItemID = ConnEnv.EnvItemID " _
        & "WHERE (ConnEnv.ProjEnvID)=" & mProjEnvID
     
        
        Set rstEnv = DB.OpenRecordset(strEnv)
        Debug.Print "rstEnv Count: " & rstEnv.RecordCount
        With rstEnv ' With S 2
            '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            Do While Not rstEnv.EOF 'Loop 2
            mEnvItemID = !EnvItemID
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++ 2 dwg +++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, nWTENVSUMHD, 1, 1, 1, 0) 'Second dwg
            TblPT(1) = TblPT(1) - nJumpENVsumHD
            Dim varAttributes1 As Variant
            varAttributes1 = blockRefObj.GetAttributes
            '{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{
                For iattr_EnvItem = LBound(varAttributes1) To UBound(varAttributes1)
                    Debug.Print varAttributes1(iattr_EnvItem).TagString
                    If varAttributes1(iattr_EnvItem).TagString = "XENVitemKey" Then
                        If IsNull(!ENVItemKey) = True Then
                            varAttributes1(iattr_EnvItem).textString = ""
                        Else
                            varAttributes1(iattr_EnvItem).textString = !ENVItemKey
                            Debug.Print varAttributes1(iattr_EnvItem).textString
                        End If
                    End If
                    Debug.Print !EnvItemName
                    If varAttributes1(iattr_EnvItem).TagString = "XENVSUMNAME" Then
                        If IsNull(!EnvItemName) = True Then
                            varAttributes1(iattr_EnvItem).textString = ""
                        Else
                            varAttributes1(iattr_EnvItem).textString = !EnvItemName
                            Debug.Print varAttributes1(iattr_EnvItem).textString
                        End If
                    End If
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    mEnvItemID = !EnvItemID
                    Debug.Print "mEnvItemID: " & mEnvItemID
                    strMat = "SELECT ConnMat.ConnMatID, ConnMat.MatID, ConnMat.EnvItemID, ConnMat.Order, Material.MatCategory, Material.MatDiv, Material.MatType, Material.MatKey, " _
                    & "Material.MatName, Material.MatThick, Material.MatR, Material.MatU, Material.MatMFGR, Material.MatNote, Material.MatLink, Material.ThermBrid, " _
                    & "Material.MatRu, Material.MatHC, Material.MatC, Material.MatF, Material.MatSHGC, Material.MatAirLeak " _
                    & "FROM Material INNER JOIN ConnMat ON Material.MatID = ConnMat.MatID " _
                    & "WHERE ((ConnMat.EnvItemID)=" & mEnvItemID & ")" _
                    & "ORDER BY ConnMat.Order "
                
                    
                    Set rstMat = DB.OpenRecordset(strMat)
                    With rstMat ' With S 3
                    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                        Do While Not rstMat.EOF 'Loop 3
                        Debug.Print "rstMat Record Count:>>>>>>>" & rstMat.RecordCount
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++ 3 dwg +++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                 (TblPT, nWTMATROW01, 1, 1, 1, 0) 'Third dwg
                        Dim varAttributes2 As Variant
                        varAttributes2 = blockRefObj.GetAttributes
                        
                        '{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{
                                For iattr_Mat = LBound(varAttributes2) To UBound(varAttributes2)
                                '************ XMAT ***************************************************************
                                    If varAttributes2(iattr_Mat).TagString = "XMAT" Then
                                        If IsNull(!MatName) = True Then
                                            varAttributes2(iattr_Mat).textString = ""
                                        Else
                                            varAttributes2(iattr_Mat).textString = !MatName
                                            Debug.Print varAttributes2(iattr_Mat).textString
                                        End If
                                    End If
                                '************ XMATR ***************************************************************
                                    If varAttributes2(iattr_Mat).TagString = "XMATR" Then
                                        If IsNull(!matr) = True Then
                                            varAttributes2(iattr_Mat).textString = ""
                                        Else
                                            varAttributes2(iattr_Mat).textString = !matr
                                            Debug.Print varAttributes2(iattr_Mat).textString
                                        End If
                                    End If
                                '************ XMATNOTE1 ***************************************************************
                                    If varAttributes2(iattr_Mat).TagString = "XMATNOTE1" Then
                                        If IsNull(!MatNote) = True Then
                                            varAttributes2(iattr_Mat).textString = ""
                                        Else
                                            varAttributes2(iattr_Mat).textString = SEPTEXT2(!MatNote, 1, 2, 12)
                                            'varAttributes1(iattr_Insp).textString = SEPTEXT2(!ENERGYWORKSCOPE, 1, 2, 40)
                                            Debug.Print varAttributes2(iattr_Mat).textString
                                        End If
                                    End If
                                '************ XMATNOTE1 ***************************************************************
                                    If varAttributes2(iattr_Mat).TagString = "XMATNOTE2" Then
                                        If IsNull(!MatNote) = True Then
                                            varAttributes2(iattr_Mat).textString = ""
                                        Else
                                            varAttributes2(iattr_Mat).textString = SEPTEXT2(!MatNote, 2, 2, 12)
                                            'varAttributes1(iattr_Insp).textString = SEPTEXT2(!ENERGYWORKSCOPE, 1, 2, 40)
                                            Debug.Print varAttributes2(iattr_Mat).textString
                                        End If
                                    End If
                                '.MoveNext
                                Next iattr_Mat
                            '}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}
                            rstMat.MoveNext
                            TblPT(1) = TblPT(1) - nJumpMATROW01
                            Loop 'Loop 3
                    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                    End With 'rstMAT ' With e 3
                    rstMat.Close
                    Set rstMat = Nothing
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
             'rstEnv.MoveNext
             'Loop 'Loop 3
                 
            Next iattr_EnvItem 'SECOND DWG END
            '}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}
            'rstEnv.MoveNext
            'Loop 'Loop 2
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++ 4 dwg +++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                 (TblPT, nWTENVSUMFOOT, 1, 1, 1, 0) 'Fourth dwg
        Dim varAttributes3 As Variant
        varAttributes3 = blockRefObj.GetAttributes
        TblPT(1) = TblPT(1) - nJumpEnvSumFoot
        For i3 = LBound(varAttributes3) To UBound(varAttributes3)
            'For iattr_Mat = LBound(varAttributes2) To UBound(varAttributes2)
            'For iattr_EnvItem = LBound(varAttributes1) To UBound(varAttributes1)
        
            '************ XENVSUMRFOOT ***************************************************************
            If varAttributes3(i3).TagString = "XENVSUMRFOOT" Then
                If IsNull(!EnvItemName) = True Then
                    varAttributes3(i3).textString = ""
                Else
                    varAttributes3(i3).textString = "Total R-Value of " & !EnvItemName
                    Debug.Print varAttributes3(i3).textString
                End If
            End If
            '************ XENVSUMUFOOT ***************************************************************
            If varAttributes3(i3).TagString = "XENVSUMUFOOT" Then
                If IsNull(!EnvItemName) = True Then
                    varAttributes3(i3).textString = ""
                Else
                    varAttributes3(i3).textString = "U-Factor of " & !EnvItemName
                    Debug.Print varAttributes3(i3).textString
                End If
            End If
            
            '************ XENVSUMR ***************************************************************
            If varAttributes3(i3).TagString = "XENVSUMR" Then
                If IsNull(!EnvItemR) = True Then
                    varAttributes3(i3).textString = ""
                Else
                    varAttributes3(i3).textString = !EnvItemR
                    Debug.Print varAttributes3(i3).textString
                End If
            End If
            '************ XENVSUMU ***************************************************************
            If varAttributes3(i3).TagString = "XENVSUMU" Then
                If IsNull(!EnvItemU) = True Then
                    varAttributes3(i3).textString = ""
                Else
                    varAttributes3(i3).textString = !EnvItemU
                    Debug.Print varAttributes3(i3).textString
                End If
            End If
            '************ XENVSUMNOTE1 ***************************************************************
            If varAttributes3(i3).TagString = "XENVSUMNOTE1" Then
                If IsNull(!EnvItemNote) = True Then
                    varAttributes3(i3).textString = ""
                Else
                    varAttributes3(i3).textString = SEPTEXT2(!EnvItemNote, 1, 4, 40)
                    Debug.Print varAttributes3(i3).textString
                End If
            End If
            '************ XENVSUMNOTE2 ***************************************************************
            If varAttributes3(i3).TagString = "XENVSUMNOTE2" Then
                If IsNull(!EnvItemNote) = True Then
                    varAttributes3(i3).textString = ""
                Else
                    varAttributes3(i3).textString = SEPTEXT2(!EnvItemNote, 2, 4, 40)
                    Debug.Print varAttributes3(i3).textString
                End If
            End If
            '************ XENVSUMNOTE3 ***************************************************************
            If varAttributes3(i3).TagString = "XENVSUMNOTE3" Then
                If IsNull(!EnvItemNote) = True Then
                    varAttributes3(i3).textString = ""
                Else
                    varAttributes3(i3).textString = SEPTEXT2(!EnvItemNote, 3, 4, 40)
                    Debug.Print varAttributes3(i3).textString
                End If
            End If
            '************ XENVSUMNOTE4 ***************************************************************
            If varAttributes3(i3).TagString = "XENVSUMNOTE4" Then
                If IsNull(!EnvItemNote) = True Then
                    varAttributes3(i3).textString = ""
                Else
                    varAttributes3(i3).textString = SEPTEXT2(!EnvItemNote, 4, 4, 40)
                    Debug.Print varAttributes3(i3).textString
                End If
            End If
        Next i3
    rstEnv.MoveNext
    Loop 'Loop 2 End 4th. dwg
    End With
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++ 5 dwg +++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        strProjENVUA = "SELECT EnvItem.EnvItemID, EnvItem.EnvItemKey, EnvItem.EnvItemName, EnvItem.EnvItemArea, EnvItem.EnvItemU, EnvItem.EnvItemNote, " _
        & "EnvItem.EnvCategory, EnvItem.EnvItemThick, EnvItem.EnvItemR, EnvItem.EnvItemThickFrac, EnvItem.ThermBrid, EnvItem.EnvItemLink, EnvItem.EnvTypeID, " _
        & "ProjEnvUA.ProjEnvID, ProjEnvUA.ProjBldgEnvID, ProjEnvUA.ProjEnvUA, ProjEnvUA.ProjEnvArea, ProjEnvUA.ProjEnvAreaUnit " _
        & "FROM (EnvItem INNER JOIN ProjEnvUA ON EnvItem.EnvItemID = ProjEnvUA.EnvItemID) INNER JOIN ConnEnv ON EnvItem.EnvItemID = ConnEnv.EnvItemID " _
        & "WHERE (ProjEnvUA.ProjEnvID)=" & mProjEnvID
        '& "WHERE (EnvItem.EnvItemID)=" & mEnvItemID
        '& "WHERE (ProjEnvUA.ProjEnvID)=" & mProjEnvID
        
        Set rstProjENVUA = DB.OpenRecordset(strProjENVUA)
        Debug.Print " rstProjENVUA : " & rstProjENVUA.RecordCount
        If rstProjENVUA.RecordCount = 0 Then
        GoTo skipProjUA
        Else
        With rstProjENVUA ' With S 4
            '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                           
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                             (TblPT, nWTPROJENVUAHD, 1, 1, 1, 0) 'Fifth dwg
            TblPT(1) = TblPT(1) - nJumpProjEnvUAHD
            Dim varAttributes4 As Variant
            varAttributes4 = blockRefObj.GetAttributes
            For i4 = LBound(varAttributes4) To UBound(varAttributes4)
                If varAttributes4(i4).TagString = "XPROJUAHDNOTE1" Then
                    If IsNull(!EnvItemNote) = True Or !EnvItemNote = "" Then
                        varAttributes4(i4).textString = ""
                    Else
                        varAttributes4(i4).textString = SEPTEXT2(!EnvItemNote, 1, 1, 40)
                        Debug.Print varAttributes4(i4).textString
                    End If
                End If
                If varAttributes4(i4).TagString = "XPROJUNITTYPE" Then
                    If IsNull(!ProjEnvAreaUnit) = True Then
                        varAttributes4(i4).textString = ""
                    Else
                        varAttributes4(i4).textString = "(" & !ProjEnvAreaUnit & ")"
                        Debug.Print varAttributes4(i4).textString
                    End If
                End If
                If varAttributes4(i4).TagString = "XPROJENVHD" Then
                    If IsNull(!EnvItemName) = True Then
                        varAttributes4(i4).textString = ""
                    Else
                        varAttributes4(i4).textString = !EnvItemName
                        Debug.Print varAttributes4(i4).textString
                    End If
                End If
            Next i4
            
            
                  
            
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++ 6 dwg +++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            Do While Not rstProjENVUA.EOF 'Loop 999
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, nWTPROJENVUArow, 1, 1, 1, 0) '6TH dwg
            TblPT(1) = TblPT(1) - nJumpProjEnvUArow
            Dim varAttributes5 As Variant
            varAttributes5 = blockRefObj.GetAttributes
            For i5 = LBound(varAttributes5) To UBound(varAttributes5)
                '************ XENVSUMNAME ***************************************************************
                If varAttributes5(i5).TagString = "XENVSUMNAME" Then
                    If IsNull(!EnvItemName) = True Then
                        varAttributes5(i5).textString = ""
                    Else
                        varAttributes5(i5).textString = !EnvItemName
                        Debug.Print varAttributes5(i5).textString
                    End If
                End If
                '************ PEU ***************************************************************
                If varAttributes5(i5).TagString = "PEU" Then
                    If IsNull(!EnvItemU) = True Then
                        varAttributes5(i5).textString = ""
                    Else
                        varAttributes5(i5).textString = !EnvItemU
                        Debug.Print varAttributes3(i5).textString
                    End If
                End If
                '************ PEA ***************************************************************
                If varAttributes5(i5).TagString = "PEA" Then
                    If IsNull(!ProjEnvArea) = True Then
                        varAttributes5(i5).textString = ""
                    Else
                        varAttributes5(i5).textString = !ProjEnvArea
                        Debug.Print varAttributes3(i5).textString
                    End If
                End If
                
                '************ UA -> ProjEnvUA ***************************************************************
                If varAttributes5(i5).TagString = "PEUA" Then
                    If IsNull(!ProjEnvUA) = True Then
                        varAttributes5(i5).textString = ""
                    Else
                        varAttributes5(i5).textString = !ProjEnvUA
                        Debug.Print varAttributes3(i5).textString
                    End If
                End If
            Next i5
            
            
            .MoveNext
            Loop 'lOOP 999
        End With ' With e 4
        End If
'skipProjUA:
    'rstProjEnvItem.MoveNext
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Debug.Print "rstProjEnvItem.RecordCount:" & rstProjEnvItem.RecordCount
Debug.Print rstProjEnvItem.Fields("projenvarea")
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++ 7 dwg +++++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
             (TblPT, nWTPROJENVUAFOOT, 1, 1, 1, 0) '7TH dwg
    TblPT(1) = TblPT(1) - nJumpProjEnvUAFOOT
    Dim varAttributes6 As Variant
    varAttributes6 = blockRefObj.GetAttributes
    For i6 = LBound(varAttributes6) To UBound(varAttributes6)
        '************ PEU ***************************************************************
        If varAttributes6(i5).TagString = "PEU" Then
            If IsNull(!EnvItemU) = True Then
                varAttributes6(i6).textString = ""
            Else
                varAttributes6(i6).textString = !EnvItemU
                Debug.Print varAttributes5(i6).textString
            End If
        End If
        '************ PEASUM ***************************************************************
        If varAttributes6(i6).TagString = "PEASUM" Then
            If IsNull(!ProjEnvArea) = True Or !ProjEnvArea = "" Then
                varAttributes6(i6).textString = ""
            Else
                varAttributes6(i6).textString = !ProjEnvArea
                Debug.Print varAttributes6(i6).textString
            End If
        End If
        '************ PEUASUM ***************************************************************
        If varAttributes6(i6).TagString = "PEUASUM" Then
            If IsNull(!ProjEnvUATotal) = True Then
                varAttributes6(i6).textString = ""
            Else
                varAttributes6(i6).textString = !ProjEnvUATotal
                Debug.Print varAttributes6(i6).textString
            End If
        End If
        '************ PEUAFIN ***************************************************************
        If varAttributes6(i6).TagString = "PEUAFIN" Then
            If IsNull(!ProjEnvU) = True Then
                varAttributes6(i6).textString = ""
            Else
                varAttributes6(i6).textString = !ProjEnvU
                Debug.Print varAttributes6(i6).textString
            End If
        End If
        '************ PEUAFINCAL ***************************************************************
        If varAttributes6(i6).TagString = "PEUAFINCAL" Then
            If IsNull(!ProjEnvU) = True Then
                varAttributes6(i6).textString = ""
            Else
                varAttributes6(i6).textString = !ProjEnvUATotal & "/" & !ProjEnvArea & "="
                Debug.Print varAttributes6(i6).textString
            End If
        End If
        '************ UTABLE ***************************************************************
        If varAttributes6(i6).TagString = "UTABLE" Then
            If IsNull(!UTABLE) = True Then
                varAttributes6(i6).textString = ""
            Else
                varAttributes6(i6).textString = !UTABLE
                Debug.Print varAttributes6(i6).textString
            End If
        End If
        '************ UTABLECITATION ***************************************************************
        If varAttributes6(i6).TagString = "UTABLECITATION" Then
            If IsNull(!CitationUtable) = True Then
                varAttributes6(i6).textString = ""
            Else
                
                varAttributes6(i6).textString = "(" & !CitationUtable & ")"
                Debug.Print varAttributes6(i6).textString
            End If
        End If
        
    Next i6
skipProjUA:
    rstProjEnvItem.MoveNext
 TblPT(0) = TblPT(0) + nShift: TblPT(1) = 0
Loop

End With


Exit_writeProjENV:
rstProjEnvItem.Close
Set rstProjEnvItem = Nothing

End Function
