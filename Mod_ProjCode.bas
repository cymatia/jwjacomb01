Attribute VB_Name = "Mod_ProjCode"

Public Sub ShowProjCode()
frmProjCode.show
End Sub

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'*********************** 9/23/18 ***************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************


Function writeProjCodeNW(xProjCodeID, xChapterID As Integer)
'On Error GoTo Err_writemtext
Dim DB As Database
Dim rstChapter, rstSection, rstSubSection As Recordset
Dim strChapter, strSection, strSubSection, CodeText As String
Dim rnChapter, rnSection, rnSubSection, rnCodePara1 As Integer
Dim iChapter, iSection, iSubSection, iCodePara1 As Integer
Dim CodeArray() As Variant
Dim StartPoint(0 To 2) As Double
Dim SectionPoint(0 To 2) As Double
Dim paraPoint(0 To 2) As Double
Dim TextPosition(0 To 2) As Double
Dim TblPT(0 To 2) As Double
Dim TblEndPT(0 To 2) As Double
Dim SizestartPoint(0 To 2) As Double
Dim width As Double
Dim mATno As Integer
Dim tmpCodeText As String


'*********************************************************************************************
'*********************************************************************************************
Set DB = OpenDatabase("\\jwja-svr-10\jdb\DB_CODE\2018-Code-01.mdb")
'*********************************************************************************************
'*********************************************************************************************
' 09/22/18
'*********************************************************************************************
'*********************************************************************************************
'mSCF= Scale Factor
If frmProjCode.OptionButton1.value Then
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
SectionPoint(0) = 0#: SectionPoint(1) = 0#: SectionPoint(2) = 0#
paraPoint(0) = 0#: paraPoint(1) = 0#: paraPoint(2) = 0#
TextPosition(0) = 0#: TextPosition(1) = 0#: TextPosition(2) = 0#
'width = 384 '32' X 12'

CodeText = ""
mChapterNo = 1
'**********************************************************************************************************
'**********************************************************************************************************

'strPCSI = "SELECT PROJCSI.PROJCSIID, PROJCSI.PROJCSIDATE, PROJCSI.PROJSPECID, " _
& "PROJCSI.PROJ_NO, MCSI.MCSIID, MCSI.MCSI, MCSI.MCSINAME " _
& "FROM MCSI INNER JOIN PROJCSI ON MCSI.MCSIID = PROJCSI.MCSIID " _
& "WHERE (((PROJCSI.PROJCSIID)=" & xProjCSIID & "));"

strChapter = "SELECT rellaw.projcodeid, rellaw.chapterid, chapter.chapterno, chapter.chaptername, chapter.codeid " _
& "FROM rellaw INNER JOIN chapter ON rellaw.chapterid = chapter.chapterID " _
& "WHERE (((rellaw.projcodeid)=" & xProjCodeID & ") AND ((rellaw.chapterid)=" & xChapterID & "));"



Set rstChapter = DB.OpenRecordset(strChapter)
rnChapter = rstChapter.RecordCount

'=============================================================
' With 01
'=============================================================
With rstChapter
'=============================================================
' Loop 01
'=============================================================

For iChapter = 1 To rnChapter
    Debug.Print "xChapterID : " & xChapterID
    '*********************** Format Start CSIName ************************************************
    If frmProjCode!mNote = True Then
    '*********************************************************************************************************
        '*********************************************************************************************************
        'text = text & Chr(10) & "\Farial.ttf;\H8.0;" & RsCSI!mCSI & " " & RsCSI!MCSIname & "\P"
        '*********************************************************************************************************
        text = text & Chr(10) & "\Farial.ttf;\H" & xScFBig & ";" & rstChapter!chapterno & " " & rstChapter!chaptername & " " & rstChapter!chapterid & "\P"
        '*********************************************************************************************************
        '*********************************************************************************************************
    Else
        'text = text & Chr(10) & RsCSI!mCSI & " " & RsCSI!MCSIname & "\P"
        'codetext = codetext & Chr(10) & rstChapter!chapterno & " " & rstChapter!chaptername & " " & rstChapter!chapterid & "\P"
        'codetext = Chr(10) & rstChapter!chapterno & " " & rstChapter!chaptername & "\P"
        CodeText = Chr(10) & "\Farial.ttf;\H" & xScFBig & ";" & rstChapter!chapterno & " - " & rstChapter!chaptername & "\P"
    End If
    '=================================================================================================================================
    '#############################
    '=================================================================================================================================
    '######### Write Chapter
    '=================================================================================================================================
    TextPosition(0) = 0#: TextPosition(1) = 0#: TextPosition(2) = 0#
    Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, CodeText)
    mtextObj.Update
    mtextObj.GetBoundingBox minText, maxText
    TextPosition(1) = minText(1) + TextPosition(1)
    TextPosition(0) = TextPosition(0) - 12
    strTextPosition = "Chapter X: " & TextPosition(0) & " Y: " & TextPosition(1)
    'Set mPositionObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, strTextPosition)
    'mPositionObj.Update
    TextPosition(0) = 0#
    '=================================================================================================================================
    '######### Write Chapter End
    '=================================================================================================================================
    '=================================================================================================================================
    
    '*********************** Format End CSIName ************************************************
    '=================================================================================================================================
    '=================================================================================================================================
    ' Related Section
    '=================================================================================================================================
    '=================================================================================================================================
    'strSection_01 = "SELECT rellaw.sectionid, section.sectionno, section.section, section.articleid, " _
    & "rellaw.relLawId, rellaw.tblName, rellaw.codeid, rellaw.chapterid, rellaw.subsectionid, rellaw.paraid, rellaw.lawno, rellaw.rellawmemo, rellaw.input, rellaw.projcodeid, rellaw.chapterno, Val([sectionno]) AS vSectionNo " _
    & "FROM [section] INNER JOIN rellaw ON section.sectionid = rellaw.sectionid " _
    & "WHERE (((rellaw.chapterid) =" & xChapterID & ") And ((rellaw.projcodeid) =" & xProjCodeid & ")) " _
    & "ORDER BY Val([sectionno]),rellaw.lawno;"
    
    strSection = "SELECT section.sectionno, section.section, section.sectionid, Val([sectionno]) AS vSectionNo, rellaw.projcodeid, rellaw.chapterid " _
    & "FROM [section] INNER JOIN rellaw ON section.sectionid = rellaw.sectionid " _
    & "GROUP BY section.sectionno, section.section, section.sectionid, Val([sectionno]), rellaw.projcodeid, rellaw.chapterid " _
    & "HAVING (((rellaw.projcodeid) = " & xProjCodeID & ") And ((rellaw.chapterid) = " & xChapterID & ")) " _
    & "ORDER BY Val([sectionno]);"


    '=================================================================================================================================
    Set rstSection = DB.OpenRecordset(strSection)
    '=============================================================
    ' With 02
    '=============================================================
    With rstSection
    .MoveFirst
    .MoveLast
    .MoveFirst
    rnSection = rstSection.RecordCount
    Debug.Print "rnSection: " & rnSection
    If rnSection > 0 Then
        Debug.Print "rnSection: " & rnSection
    End If
    '=============================================================
    ' Loop 02 Section
    '=============================================================
    For iSection = 1 To rnSection
        If frmProjCode!mNote = True Then
            
            'text = text & "\P" & Chr(9) & RsPart!MPART & "\P"
            Debug.Print rstSection!section
            CodeText = CodeText
        Else
            'codetext = codetext & "\P" & "Section " & rstSection!sectionno & Chr(9) & rstSection!section & "\P"
            'Set MTextObj = ThisDrawing.ModelSpace.AddMText(corner, width, text)
            CodeText = "Section " & rstSection!sectionno & Chr(9) & rstSection!section & "\P"
        End If
        '=================================================================================================================================
        '######### Write Section
        '=================================================================================================================================
        Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, CodeText)
        If frmProjCode!mNote = True Then
            CreateARstyle
            mtextObj.StyleName = "ar"
            mtextObj.height = xScF
        Else
            CreateCNstyle
            mtextObj.StyleName = "cn"
            mtextObj.height = xScF
        End If
        mtextObj.Update
        mtextObj.GetBoundingBox minText, maxText
        'TextPosition(1) = minText(1) + TextPosition(1)
        TextPosition(1) = minText(1)
        strTextPosition = "Section X: " & TextPosition(0) & "Y: " & TextPosition(1) + 8
        'Set mPositionObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, strTextPosition)
        'mPositionObj.Update
        '=================================================================================================================================
        '######### Write Section End
        '=================================================================================================================================
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        xSectionId = rstSection!sectionid
        
        '=================================================================================================================================
        '=================================================================================================================================
        ' Related CodePara
        '=================================================================================================================================
        '=================================================================================================================================
        CodeText = ""
        strCodePara = " SELECT rellaw.lawno, rellaw.relLawId, rellaw.tblName, rellaw.codeid, rellaw.chapterid, rellaw.sectionid, rellaw.subsectionid, " _
        & "rellaw.paraid, rellaw.rellawmemo, rellaw.input, rellaw.projcodeid, rellaw.sectionid " _
        & "FROM rellaw " _
        & "WHERE (((rellaw.projcodeid) = " & xProjCodeID & ") And ((rellaw.sectionid) = " & xSectionId & ")) " _
        & "ORDER BY rellaw.lawno;"
        Set rstCodePara = DB.OpenRecordset(strCodePara)
        '=============================================================
        ' Loop 03 Para
        '=============================================================
        Do While Not rstCodePara.EOF
        Debug.Print rstCodePara!rellawmemo
        'codetext = codetext & "\P" & rstCodePara!lawno & Chr(9) & rstCodePara!tblname & Chr(9) & rstCodePara!paraid & "\P"
        Select Case rstCodePara!tblName
        '+++++++++++++++++++++++++++++++++++++++++++++
        '********************** Problem Here!!! ******************
        
        Case "subsection"
        
            'strCodeText = "SELECT Val([subsectionno]) AS Expr1, subsection.lawno, rellaw.projcodeid, rellaw.tblname, rellaw.paraid, subsection.sectionid, subsection.subsection, subsection.subsectionid " _
            & "FROM subsection INNER JOIN rellaw ON subsection.subsectionid = rellaw.subsectionid " _
            & "GROUP BY Val([subsectionno]), subsection.lawno, rellaw.projcodeid, rellaw.tblname, rellaw.paraid, subsection.sectionid, subsection.subsection, subsection.subsectionid " _
            & "HAVING (((rellaw.projcodeid) = " & xProjCodeID & ") And ((subsection.sectionid) = " & xSectionId & ")) " _
            & "ORDER BY Val([subsectionno]), subsection.lawno;"
            '****************Corrected below 18/11/11
            'strCodeText_old = "SELECT subsection.lawno, subsection.subsectionid, subsection.subsectionno, subsection.subsection, subsection.sectionid " _
            & "FROM subsection " _
            & "GROUP BY subsection.lawno, subsection.subsectionid, subsection.subsectionno, subsection.subsection, subsection.sectionid " _
            & "HAVING (((subsection.subsectionid) = " & rstCodePara!paraid & ")) " _
            & "ORDER BY subsection.lawno;"
            
            strCodeText = "SELECT subsection.lawno, subsection.subsectionid, subsection.subsectionno, subsection.subsection, subsection.sectionid " _
            & "FROM subsection " _
            & "where (((subsection.subsectionid) = " & rstCodePara!paraid & ")) " _
            & "ORDER BY subsection.lawno;"
            'test 551
            
            Set rstCodeText = DB.OpenRecordset(strCodeText)
            rnrstCodeText = rstCodeText.RecordCount
            If (rnrstCodeText) = 0 Then
            
            GoTo EndselectPara
            Else
            tmpCodeText = rstCodeText!subsection
            End If
            Debug.Print rstCodeText!subsection
            tmpException = writeException(xProjCodeID, "subsection", rstCodeText!subsectionid)
            Debug.Print tmpCodeText
            tmpMemo = ""
            Debug.Print rstCodePara!rellawmemo
            If IsNull(rstCodePara!rellawmemo) = False Then

            tmpMemo = rstCodePara!rellawmemo
            End If
            
            rstCodeText.Close
            Set rstCodeText = Nothing
        '+++++++++++++++++++++++++++++++++++++++++++++
        Case "codepara1"
            'codetext = codetext & "\P" & "xxxPara1" & "\P"
            strCodeText = "SELECT codepara1.codepara1id, codepara1.codepara1no, codepara1.lawno, codepara1.codepara1 " _
            & "FROM codepara1 " _
            & "WHERE (((codepara1.codepara1id) =" & rstCodePara!paraid & ")) " _
            & "ORDER BY codepara1.lawno;"
            Set rstCodeText = DB.OpenRecordset(strCodeText)
            'Do While Not rstCodeText.EOF
            tmpCodeText = rstCodeText!codepara1
            tmpMemo = ""
            If IsNull(rstCodePara!rellawmemo) = False Then
            tmpMemo = rstCodePara!rellawmemo
            End If
            
            rstCodeText.Close
            Set rstCodeText = Nothing
            'Loop
        Case "codepara2"
            'codetext = codetext & "\P" & "xxxPara2" & "\P"
            strCodeText = "SELECT codepara2.codepara2id, codepara2.codepara2no, codepara2.lawno, codepara2.codepara2 " _
            & "FROM codepara2 " _
            & "WHERE (((codepara2.codepara2id) =" & rstCodePara!paraid & ")) " _
            & "ORDER BY codepara2.lawno;"
            Set rstCodeText = DB.OpenRecordset(strCodeText)
            'Do While Not rstCodeText.EOF
            tmpCodeText = rstCodeText!codepara2
            tmpMemo = ""
    
            If IsNull(rstCodePara!rellawmemo) = False Then
            tmpMemo = rstCodePara!rellawmemo
            End If
            rstCodeText.Close
            Set rstCodeText = Nothing
        
        Case "codepara3"
            'codetext = codetext & "\P" & "xxxPara3" & "\P"
            strCodeText = "SELECT codepara3.codepara3id, codepara3.codepara3no, codepara3.lawno, codepara3.codepara3 " _
            & "FROM codepara3 " _
            & "WHERE (((codepara3.codepara3id) =" & rstCodePara!paraid & ")) " _
            & "ORDER BY codepara3.lawno;"
            Set rstCodeText = DB.OpenRecordset(strCodeText)
            'Do While Not rstCodeText.EOF
            tmpCodeText = rstCodeText!codepara3
            tmpMemo = ""
            If IsNull(rstCodePara!rellawmemo) = False Then
            tmpMemo = rstCodePara!rellawmemo
            End If
            rstCodeText.Close
            Set rstCodeText = Nothing
        
        Case "codeepara4"
            'codetext = codetext & "\P" & "xxxPara4" & "\P"
            strCodeText = "SELECT codepara4.codepara4id, codepara4.codepara4no, codepara4.lawno, codepara4.codepara4 " _
            & "FROM codepara4 " _
            & "WHERE (((codepara4.codepara4id) =" & rstCodePara!paraid & ")) " _
            & "ORDER BY codepara4.lawno;"
            Set rstCodeText = DB.OpenRecordset(strCodeText)
            'Do While Not rstCodeText.EOF
            tmpCodeText = rstCodeText!codepara4
            tmpMemo = ""
            If IsNull(rstCodePara!rellawmemo) = False Then
            tmpMemo = rstCodePara!rellawmemo
            End If
            rstCodeText.Close
            Set rstCodeText = Nothing
        Case "exception"
            'codetext = codetext & "\P" & "xxxPara4" & "\P"
             strCodeText = "SELECT exception.exceptionid, exception.exceptionno, exception.exception, exception.codepara1id, exception.codepara2id, exception.lawno, exception.subsectionid, exception.sectionid, exception.tblname, exception.paraid, exception.codeid " _
            & "FROM [exception] " _
            & "WHERE (((exception.exceptionid) =" & rstCodePara!paraid & ")) " _

            
            Set rstCodeText = DB.OpenRecordset(strCodeText)
            'Do While Not rstCodeText.EOF
            tmpCodeText = rstCodeText!exception
            tmpMemo = ""
            If IsNull(rstCodePara!rellawmemo) = False Then
            tmpMemo = rstCodePara!rellawmemo
            End If
            rstCodeText.Close
            Set rstCodeText = Nothing
            'exception
EndselectPara:
        End Select
    '=================================================================================================================================
    '######### Write Para
    '=================================================================================================================================
    If (tmpMemo) = "" Then
        CodeText = CodeText & "\P" & rstCodePara!lawno & Chr(9) & tmpCodeText & "\P"
    Else
        'codetext = codetext & "\P" & rstCodePara!lawno & Chr(9) & tmpCodeText & "\P" & "\P" & BigText(tmpMemo) & "\P"
        CodeText = CodeText & "\P" & rstCodePara!lawno & Chr(9) & tmpCodeText & "\P"
        'TblPT(0) = mTab + TextPosition(0)
        'TblEndPT(0) = TblPT(0) + 32
        'ThisDrawing.ModelSpace.AddLine TblPT, TblEndPT
        'codetext = codetext & "\P" & BigText(tmpMemo) & "\P"
        
    End If
    'tmpMemo = ""
    
    'startPoint(0) = 0#: startPoint(1) = 0#: startPoint(2) = 0#
    '1440 x 0.5 = 720
    mTab = 12
    TextPosition(0) = mTab + TextPosition(0)
    Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, CodeText)
    If frmProjCode!mNote = True Then
        CreateARstyle
        mtextObj.StyleName = "ar"
        mtextObj.height = xScF
    Else
        CreateCNstyle
        mtextObj.StyleName = "cn"
        mtextObj.height = xScF
    End If

    mtextObj.Update
    mtextObj.GetBoundingBox minText, maxText
    TextPosition(1) = minText(1)
    If (tmpMemo) = "" Then
    Else
        'width = 384 '32' X 12'
        TblPT(0) = mTab + mTab + TextPosition(0)
        TblPT(1) = TextPosition(1)
        TblEndPT(0) = TblPT(0) + (width - TblPT(0))
        TblEndPT(1) = TextPosition(1)
        ThisDrawing.ModelSpace.AddLine TblPT, TblEndPT
        
        codeMemo = "\P" & BigText("Project Note: " & tmpMemo, xScFBig) & "\P"
         TextPosition(0) = mTab + mTab + TextPosition(0)
        Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, codeMemo)
        TextPosition(0) = TextPosition(0) - mTab - mTab
        mtextObj.StyleName = "a"
        mtextObj.height = xScF
        mtextObj.Update
        
       
        mtextObj.GetBoundingBox minText, maxText
        TextPosition(1) = minText(1)
        TblPT(0) = mTab + mTab + TextPosition(0)
        TblPT(1) = TextPosition(1)
        TblEndPT(0) = TblPT(0) + (width - TblPT(0))
        TblEndPT(1) = TextPosition(1)
        ThisDrawing.ModelSpace.AddLine TblPT, TblEndPT
    End If
    tmpMemo = ""

    
    
    TextPosition(0) = 0#
    'TextPosition(1) = minText(1) + TextPosition(1)
    TextPosition(1) = minText(1)
    
    strTextPosition = "Section X: " & TextPosition(0) & "Y: " & TextPosition(1)
    'Set mPositionObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, strTextPosition)
    'mPositionObj.Update
    CodeText = ""
    '=================================================================================================================================
    '######### Write Para End
    '=================================================================================================================================
nextiPara:
        rstCodePara.MoveNext
        Loop
        '=============================================================
        ' Loop 03 Para End
        '=============================================================
End03:
               
        rstCodePara.Close
        Set rstCodePara = Nothing
    '-----------------------------------------------------------------------------------------------------
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
'=============================================================
' Loop 02 Section end
'=============================================================
    rstSection.MoveNext
    Next iSection
    End With
    rstSection.Close
'rstChapter.MoveNext
'=============================================================
' Loop 01 end
'=============================================================
Next iChapter
                'rstSubSection.Close
               ' Set rstSubSection = Nothing
                'rstCodePara1.Close
                'Set rstCodePara1 = Nothing
                'rstCodePara2.Close
                'Set rstCodePara2 = Nothing
                'rstCodePara3.Close
                'Set rstCodePara3 = Nothing
                
                
'Set mtextObj = ThisDrawing.ModelSpace.AddMText(startPoint, width, codetext)
'mtextObj.Update
End With
'=============================================================
'With 01 End
'=============================================================
rstChapter.Close

Set rstChapter = Nothing
DB.Close
'mtextObj.GetBoundingBox minExt, maxExt
'SizestartPoint(0) = 0#: SizestartPoint(1) = minExt(1) + 0#: SizestartPoint(2) = 0#
'textobjSize = "minExt: " & minExt(1) & " maxExt: " & maxExt(1)
'Set mtextobjsize = ThisDrawing.ModelSpace.AddMText(SizestartPoint, width, textobjSize)
'mtextobjsize.Update

Debug.Print rnPart & rnAt & rnPara1
       
End Function
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
Function writeProjCode(xProjCodeID, xChapterID As Integer)
'On Error GoTo Err_writemtext
Dim DB As Database
Dim rstChapter, rstSection, rstSubSection As Recordset
Dim strChapter, strSection, strSubSection, CodeText As String
Dim rnChapter, rnSection, rnSubSection, rnCodePara1 As Integer
Dim iChapter, iSection, iSubSection, iCodePara1 As Integer
Dim CodeArray() As Variant
Dim StartPoint(0 To 2) As Double
Dim SectionPoint(0 To 2) As Double
Dim paraPoint(0 To 2) As Double
Dim TextPosition(0 To 2) As Double
Dim TblPT(0 To 2) As Double
Dim TblEndPT(0 To 2) As Double
Dim SizestartPoint(0 To 2) As Double
Dim width As Double
Dim mATno As Integer


'**********************************************************************************************************
'**********************************************************************************************************
Set DB = OpenDatabase("\\jwja-svr-10\jdb\DB_CODE\2018-Code-01.mdb")


'*********************************************************************************************
'*********************************************************************************************
' 09/22/18
'*********************************************************************************************
'*********************************************************************************************
'mSCF= Scale Factor
If frmProjCode.OptionButton1.value Then
    mSCF = 48 'mSCF= Scale Factor
    width = mSCF * 8 '48 x 8 = 384
ElseIf frmProjCode.OptionButton2.value Then
    mSCF = 96
    width = mSCF * 8 '96 x 8 = 768
ElseIf frmProjCode.OptionButton3.value Then
    mSCF = 64
    width = mSCF * 8 '64 x 8 = 512
End If

xScF8 = mSCF / 12 * 2
xScF6 = mSCF / 12 * 1.5
xScF4 = mSCF / 12 * 1
    'mSCF = 48 'mSCF= Scale Factor
    'width = mSCF * 8 '48 x 8 = 384
    'width = 1440 * 6
    If frmProjCode!mNote = True Then
        CreateARstyle
        mtextObj.StyleName = "ar"
        mtextObj.height = 4
    Else
        CreateCNstyle
        mtextObj.StyleName = "cn"
        mtextObj.height = 4
    End If
'*********************************************************************************************
'*********************************************************************************************
'startPoint(0) = 0#: startPoint(1) = 0#: startPoint(2) = 0#
SectionPoint(0) = 0#: SectionPoint(1) = 0#: SectionPoint(2) = 0#
paraPoint(0) = 0#: paraPoint(1) = 0#: paraPoint(2) = 0#
TextPosition(0) = 0#: TextPosition(1) = 0#: TextPosition(2) = 0#
width = 384 '32' X 12'

CodeText = ""
mChapterNo = 1
'**********************************************************************************************************
'**********************************************************************************************************

'strPCSI = "SELECT PROJCSI.PROJCSIID, PROJCSI.PROJCSIDATE, PROJCSI.PROJSPECID, " _
& "PROJCSI.PROJ_NO, MCSI.MCSIID, MCSI.MCSI, MCSI.MCSINAME " _
& "FROM MCSI INNER JOIN PROJCSI ON MCSI.MCSIID = PROJCSI.MCSIID " _
& "WHERE (((PROJCSI.PROJCSIID)=" & xProjCSIID & "));"

strChapter = "SELECT rellaw.projcodeid, rellaw.chapterid, chapter.chapterno, chapter.chaptername, chapter.codeid " _
& "FROM rellaw INNER JOIN chapter ON rellaw.chapterid = chapter.chapterID " _
& "WHERE (((rellaw.projcodeid)=" & xProjCodeID & ") AND ((rellaw.chapterid)=" & xChapterID & "));"



Set rstChapter = DB.OpenRecordset(strChapter)
rnChapter = rstChapter.RecordCount

'=============================================================
' With 01
'=============================================================
With rstChapter
'=============================================================
' Loop 01
'=============================================================

For iChapter = 1 To rnChapter
    Debug.Print "xChapterID : " & xChapterID
    '*********************** Format Start CSIName ************************************************
    If frmProjCode!mNote = True Then
    '*********************************************************************************************************
        '*********************************************************************************************************
        'text = text & Chr(10) & "\Farial.ttf;\H8.0;" & RsCSI!mCSI & " " & RsCSI!MCSIname & "\P"
        '*********************************************************************************************************
        text = text & Chr(10) & "\Farial.ttf;\H" & xScF8 & ";" & rstChapter!chapterno & " " & rstChapter!chaptername & " " & rstChapter!chapterid & "\P"
        '*********************************************************************************************************
        '*********************************************************************************************************
    Else
        'text = text & Chr(10) & RsCSI!mCSI & " " & RsCSI!MCSIname & "\P"
        'codetext = codetext & Chr(10) & rstChapter!chapterno & " " & rstChapter!chaptername & " " & rstChapter!chapterid & "\P"
        CodeText = Chr(10) & rstChapter!chapterno & " " & rstChapter!chaptername & " <" & rstChapter!chapterid & "> \P"
    End If
    '=================================================================================================================================
    '#############################
    '=================================================================================================================================
    '######### Write Chapter
    '=================================================================================================================================
    TextPosition(0) = 0#: TextPosition(1) = 0#: TextPosition(2) = 0#
    Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, CodeText)
    mtextObj.Update
    mtextObj.GetBoundingBox minText, maxText
    TextPosition(1) = minText(1) + TextPosition(1)
    TextPosition(0) = TextPosition(0) - 12
    strTextPosition = "Chapter X: " & TextPosition(0) & " Y: " & TextPosition(1)
    Set mPositionObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, strTextPosition)
    mPositionObj.Update
    TextPosition(0) = 0#
    '=================================================================================================================================
    '######### Write Chapter End
    '=================================================================================================================================
    '=================================================================================================================================
    
    '*********************** Format End CSIName ************************************************
    '=================================================================================================================================
    '=================================================================================================================================
    ' Related Section
    '=================================================================================================================================
    '=================================================================================================================================
    'strSection_01 = "SELECT rellaw.sectionid, section.sectionno, section.section, section.articleid, " _
    & "rellaw.relLawId, rellaw.tblName, rellaw.codeid, rellaw.chapterid, rellaw.subsectionid, rellaw.paraid, rellaw.lawno, rellaw.rellawmemo, rellaw.input, rellaw.projcodeid, rellaw.chapterno, Val([sectionno]) AS vSectionNo " _
    & "FROM [section] INNER JOIN rellaw ON section.sectionid = rellaw.sectionid " _
    & "WHERE (((rellaw.chapterid) =" & xChapterID & ") And ((rellaw.projcodeid) =" & xProjCodeid & ")) " _
    & "ORDER BY Val([sectionno]),rellaw.lawno;"
    
    strSection = "SELECT section.sectionno, section.section, section.sectionid, Val([sectionno]) AS vSectionNo, rellaw.projcodeid, rellaw.chapterid " _
    & "FROM [section] INNER JOIN rellaw ON section.sectionid = rellaw.sectionid " _
    & "GROUP BY section.sectionno, section.section, section.sectionid, Val([sectionno]), rellaw.projcodeid, rellaw.chapterid " _
    & "HAVING (((rellaw.projcodeid) = " & xProjCodeID & ") And ((rellaw.chapterid) = " & xChapterID & ")) " _
    & "ORDER BY Val([sectionno]);"


    '=================================================================================================================================
    Set rstSection = DB.OpenRecordset(strSection)
    '=============================================================
    ' With 02
    '=============================================================
    With rstSection
    .MoveFirst
    .MoveLast
    .MoveFirst
    rnSection = rstSection.RecordCount
    Debug.Print "rnSection: " & rnSection
    If rnSection > 0 Then
        Debug.Print "rnSection: " & rnSection
    End If
    '=============================================================
    ' Loop 02 Section
    '=============================================================
    For iSection = 1 To rnSection
        If frmProjCode!mNote = True Then
            
            'text = text & "\P" & Chr(9) & RsPart!MPART & "\P"
            Debug.Print rstSection!section
            CodeText = CodeText
        Else
            'codetext = codetext & "\P" & "Section " & rstSection!sectionno & Chr(9) & rstSection!section & "\P"
            'Set MTextObj = ThisDrawing.ModelSpace.AddMText(corner, width, text)
            CodeText = "Section " & rstSection!sectionno & Chr(9) & rstSection!section & "\P"
        End If
        '=================================================================================================================================
        '######### Write Section
        '=================================================================================================================================
        Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, CodeText)
        mtextObj.Update
        mtextObj.GetBoundingBox minText, maxText
        'TextPosition(1) = minText(1) + TextPosition(1)
        TextPosition(1) = minText(1)
        strTextPosition = "Section X: " & TextPosition(0) & "Y: " & TextPosition(1) + 8
        Set mPositionObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, strTextPosition)
        mPositionObj.Update
        '=================================================================================================================================
        '######### Write Section End
        '=================================================================================================================================
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        xSectionId = rstSection!sectionid
        
        '=================================================================================================================================
        '=================================================================================================================================
        ' Related CodePara
        '=================================================================================================================================
        '=================================================================================================================================
        CodeText = ""
        strCodePara = " SELECT rellaw.lawno, rellaw.relLawId, rellaw.tblName, rellaw.codeid, rellaw.chapterid, rellaw.sectionid, rellaw.subsectionid, " _
        & "rellaw.paraid, rellaw.rellawmemo, rellaw.input, rellaw.projcodeid, rellaw.sectionid " _
        & "FROM rellaw " _
        & "WHERE (((rellaw.projcodeid) = " & xProjCodeID & ") And ((rellaw.sectionid) = " & xSectionId & ")) " _
        & "ORDER BY rellaw.lawno;"
        mTab = 12
        TextPosition(0) = mTab + TextPosition(0)
        Set rstCodePara = DB.OpenRecordset(strCodePara)
        '=============================================================
        ' Loop 03 Para
        '=============================================================
        Do While Not rstCodePara.EOF
        
        'codetext = codetext & "\P" & rstCodePara!lawno & Chr(9) & rstCodePara!tblname & Chr(9) & rstCodePara!paraid & "\P"
        Select Case rstCodePara!tblName
        '+++++++++++++++++++++++++++++++++++++++++++++
        Case "subsection"
        
            strCodeText = "SELECT Val([subsectionno]) AS Expr1, subsection.lawno, rellaw.projcodeid, rellaw.tblname, rellaw.paraid, subsection.sectionid, subsection.subsection, subsection.subsectionid " _
            & "FROM subsection INNER JOIN rellaw ON subsection.subsectionid = rellaw.subsectionid " _
            & "GROUP BY Val([subsectionno]), subsection.lawno, rellaw.projcodeid, rellaw.tblname, rellaw.paraid, subsection.sectionid, subsection.subsection, subsection.subsectionid " _
            & "HAVING (((rellaw.projcodeid) = 39) And ((subsection.sectionid) = 269)) " _
            & "ORDER BY Val([subsectionno]), subsection.lawno;"
            Set rstCodeText = DB.OpenRecordset(strCodeText)
            tmpCodeText = rstCodeText!subsection
            tmpException = writeException(xProjCodeID, rstCodeText!tblName, rstCodeText!paraid)
            tmpMemo = ""
            If IsNull(rstCodePara!rellawmemo) = False Then
           

            tmpMemo = rstCodePara!rellawmemo
            End If
            
            rstCodeText.Close
            Set rstCodeText = Nothing
        '+++++++++++++++++++++++++++++++++++++++++++++
        Case "codepara1"
            'codetext = codetext & "\P" & "xxxPara1" & "\P"
            strCodeText = "SELECT codepara1.codepara1id, codepara1.codepara1no, codepara1.lawno, codepara1.codepara1 " _
            & "FROM codepara1 " _
            & "WHERE (((codepara1.codepara1id) =" & rstCodePara!paraid & ")) " _
            & "ORDER BY codepara1.lawno;"
            Set rstCodeText = DB.OpenRecordset(strCodeText)
            'Do While Not rstCodeText.EOF
            tmpCodeText = rstCodeText!codepara1
            tmpMemo = ""
            If IsNull(rstCodePara!rellawmemo) = False Then
            tmpMemo = rstCodePara!rellawmemo
            End If
            
            rstCodeText.Close
            Set rstCodeText = Nothing
            'Loop
        Case "codepara2"
            'codetext = codetext & "\P" & "xxxPara2" & "\P"
            strCodeText = "SELECT codepara2.codepara2id, codepara2.codepara2no, codepara2.lawno, codepara2.codepara2 " _
            & "FROM codepara2 " _
            & "WHERE (((codepara2.codepara2id) =" & rstCodePara!paraid & ")) " _
            & "ORDER BY codepara2.lawno;"
            Set rstCodeText = DB.OpenRecordset(strCodeText)
            'Do While Not rstCodeText.EOF
            tmpCodeText = rstCodeText!codepara2
            tmpMemo = ""
    
            If IsNull(rstCodePara!rellawmemo) = False Then
            tmpMemo = rstCodePara!rellawmemo
            End If
            rstCodeText.Close
            Set rstCodeText = Nothing
        
        Case "codepara3"
            'codetext = codetext & "\P" & "xxxPara3" & "\P"
            strCodeText = "SELECT codepara3.codepara3id, codepara3.codepara3no, codepara3.lawno, codepara3.codepara3 " _
            & "FROM codepara3 " _
            & "WHERE (((codepara3.codepara3id) =" & rstCodePara!paraid & ")) " _
            & "ORDER BY codepara3.lawno;"
            Set rstCodeText = DB.OpenRecordset(strCodeText)
            'Do While Not rstCodeText.EOF
            tmpCodeText = rstCodeText!codepara3
            tmpMemo = ""
            If IsNull(rstCodePara!rellawmemo) = False Then
            tmpMemo = rstCodePara!rellawmemo
            End If
            rstCodeText.Close
            Set rstCodeText = Nothing
        
        Case "codeepara4"
            'codetext = codetext & "\P" & "xxxPara4" & "\P"
            strCodeText = "SELECT codepara4.codepara4id, codepara4.codepara4no, codepara4.lawno, codepara4.codepara4 " _
            & "FROM codepara4 " _
            & "WHERE (((codepara4.codepara4id) =" & rstCodePara!paraid & ")) " _
            & "ORDER BY codepara4.lawno;"
            Set rstCodeText = DB.OpenRecordset(strCodeText)
            'Do While Not rstCodeText.EOF
            tmpCodeText = rstCodeText!codepara4
            tmpMemo = ""
            If IsNull(rstCodePara!rellawmemo) = False Then
            tmpMemo = rstCodePara!rellawmemo
            End If
            rstCodeText.Close
            Set rstCodeText = Nothing
        Case "exception"
            'codetext = codetext & "\P" & "xxxPara4" & "\P"
             strCodeText = "SELECT exception.exceptionid, exception.exceptionno, exception.exception, exception.codepara1id, exception.codepara2id, exception.lawno, exception.subsectionid, exception.sectionid, exception.tblname, exception.paraid, exception.codeid " _
            & "FROM [exception] " _
            & "WHERE (((exception.exceptionid) =" & rstCodePara!paraid & ")) " _

            
            Set rstCodeText = DB.OpenRecordset(strCodeText)
            'Do While Not rstCodeText.EOF
            tmpCodeText = rstCodeText!exception
            tmpMemo = ""
            If IsNull(rstCodePara!rellawmemo) = False Then
            tmpMemo = rstCodePara!rellawmemo
            End If
            rstCodeText.Close
            Set rstCodeText = Nothing
            'exception
        End Select
        
nextiPara:
    If (tmpMemo) = "" Then
        CodeText = CodeText & "\P" & rstCodePara!lawno & Chr(9) & tmpCodeText & "\P"
    Else
        'codetext = codetext & "\P" & rstCodePara!lawno & Chr(9) & tmpCodeText & "\P" & "\P" & BigText(tmpMemo) & "\P"
        CodeText = CodeText & "\P" & rstCodePara!lawno & Chr(9) & tmpCodeText & "\P"
        TblPT(0) = mTab + TextPosition(0)
        TblEndPT(0) = TblPT(0) + 32
        ThisDrawing.ModelSpace.AddLine TblPT, TblEndPT
        CodeText = CodeText & "\P" & BigText(tmpMemo) & "\P"
        
    End If
    tmpMemo = ""
        rstCodePara.MoveNext
        Loop
        '=============================================================
        ' Loop 03 Para End
        '=============================================================
End03:
               
        rstCodePara.Close
        Set rstCodePara = Nothing
    '-----------------------------------------------------------------------------------------------------
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    '=================================================================================================================================
    '######### Write Para
    '=================================================================================================================================
    'startPoint(0) = 0#: startPoint(1) = 0#: startPoint(2) = 0#
    '1440 x 0.5 = 720
    
    Set mtextObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, CodeText)
    mtextObj.Update
    mtextObj.GetBoundingBox minText, maxText
    TextPosition(0) = 0#
    'TextPosition(1) = minText(1) + TextPosition(1)
    TextPosition(1) = minText(1)
    strTextPosition = "Section X: " & TextPosition(0) & "Y: " & TextPosition(1)
    Set mPositionObj = ThisDrawing.ModelSpace.AddMText(TextPosition, width, strTextPosition)
    mPositionObj.Update
    CodeText = ""
    '=================================================================================================================================
    '######### Write Para End
    '=================================================================================================================================
'=============================================================
' Loop 02 Section end
'=============================================================
    rstSection.MoveNext
    Next iSection
    End With
    rstSection.Close
'rstChapter.MoveNext
'=============================================================
' Loop 01 end
'=============================================================
Next iChapter
                'rstSubSection.Close
               ' Set rstSubSection = Nothing
                'rstCodePara1.Close
                'Set rstCodePara1 = Nothing
                'rstCodePara2.Close
                'Set rstCodePara2 = Nothing
                'rstCodePara3.Close
                'Set rstCodePara3 = Nothing
                
                
'Set mtextObj = ThisDrawing.ModelSpace.AddMText(startPoint, width, codetext)
'mtextObj.Update
End With
'=============================================================
'With 01 End
'=============================================================
rstChapter.Close

Set rstChapter = Nothing
DB.Close
'mtextObj.GetBoundingBox minExt, maxExt
'SizestartPoint(0) = 0#: SizestartPoint(1) = minExt(1) + 0#: SizestartPoint(2) = 0#
'textobjSize = "minExt: " & minExt(1) & " maxExt: " & maxExt(1)
'Set mtextobjsize = ThisDrawing.ModelSpace.AddMText(SizestartPoint, width, textobjSize)
'mtextobjsize.Update

Debug.Print rnPart & rnAt & rnPara1
       
End Function
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

Function writeException(xProjCodeID, strTblName, xVarID As Variant)
Dim DB As Database
Dim rstRelException As Recordset
Dim strException As String
Dim rnChapter, rnSection, rnSubSection, rnCodePara1 As Integer
writeException = ""
Set DB = OpenDatabase("\\jwja-svr-10\jdb\DB_CODE\2018-Code-01.mdb")
strException = "SELECT rellaw.lawno, rellaw.relLawId, rellaw.tblName, rellaw.codeid, rellaw.chapterid, rellaw.sectionid, " _
& "rellaw.subsectionid, rellaw.rellawmemo, rellaw.paraid, rellaw.input, rellaw.projcodeid, rellaw.codepara2id, exception.tblname, " _
& "exception.exception, exception.subsectionid, exception.sectionid " _
& "FROM rellaw INNER JOIN [exception] ON rellaw.paraid = exception.exceptionid " _
& "WHERE (((rellaw.tblName)='exception') AND ((rellaw.projcodeid)=" & xProjCodeID & ") AND ((exception.tblname)='" & strTblName & "') AND ((exception.subsectionid)=" & xVarID & ")) " _
& "ORDER BY rellaw.lawno;"


Set rstRelException = DB.OpenRecordset(strException)
With rstRelException
rnRelException = rstRelException.RecordCount
If IsNull(rnRelException) = True Then
Exit Function
Else
For iRelException = 1 To rnRelException
writeException = writeException & "\P" & rstRelException!lawno & Chr(9) & rstRelException!exception & "\P"
Debug.Print writeException
Next iRelException
End If
End With
End Function

Sub CreateCNstyle_spec()
Dim newText As AcadTextStyle
Dim blnFoundText As Boolean
Dim blnFoundTitle As Boolean
Dim blnFoundSub As Boolean
   
'reset blnFound
blnFoundText = False
blnFoundTitle = False
blnFoundSub = False

    For i = 0 To ThisDrawing.TextStyles.Count - 1  'make the styles for the form
        Debug.Print ThisDrawing.TextStyles(i).Name
        If ThisDrawing.TextStyles(i).Name = "CN" Then
            blnFoundText = True
            Exit Sub
        End If
        
    Next i

Set newText = ThisDrawing.TextStyles.Add("CN")

typeface = "Courier new"
newText.SetFont typeface, Bold, Italic, Charset, PitchandFamily

End Sub

