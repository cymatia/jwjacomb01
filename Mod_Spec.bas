Attribute VB_Name = "Mod_Spec"

 Public Function Num2Frac(ByVal x As Double) As String
      Dim Fixed As Double, Temp As String
        x = Abs(x)
        Fixed = Int(x)
        If Fixed > 0 Then
          Temp = CStr(Fixed)
        End If
        Select Case x - Fixed
          Case 0.1 To 0.145
            Temp = Temp + " 1/8"
          Case 0.145 To 0.182
            Temp = Temp + " 1/6"
          Case 0.182 To 0.225
            Temp = Temp + " 1/5"
          Case 0.225 To 0.29
            Temp = Temp + " 1/4"
          Case 0.29 To 0.35
            Temp = Temp + " 1/3"
          Case 0.35 To 0.3875
            Temp = Temp + " 3/8"
          Case 0.3875 To 0.45
            Temp = Temp + " 2/5"
          Case 0.45 To 0.55
            Temp = Temp + " 1/2"
          Case 0.55 To 0.6175
            Temp = Temp + " 3/5"
          Case 0.6175 To 0.64
            Temp = Temp + " 5/8"
          Case 0.64 To 0.7
            Temp = Temp + " 2/3"
          Case 0.7 To 0.775
            Temp = Temp + " 3/4"
          Case 0.775 To 0.8375
            Temp = Temp + " 4/5"
          Case 0.8735 To 0.91
            Temp = Temp + " 7/8"
          Case Is > 0.91
            Temp = CStr(Int(x) + 1)
        End Select
        Num2Frac = Temp
      End Function

Public Sub jSpec() '10/13/13
frmSpec.show
End Sub


Function printspec(xProjCSIID As Integer)
Dim DB As Database
Dim RsPart, RsPat, RsPpara1 As Recordset
Dim strPpart, strPat, strPpara1 As String
Dim rnPart, rnAt, rnPara1 As Integer
Dim specArray() As Variant

'**********************************************************************************************************
'**********************************************************************************************************
'H_H_Jspec05_CURRENT
'Set DB = OpenDatabase("H:\db\db_SPEC\H_H_JSPEC04_02.mdb") '10/5/14
'Set DB = OpenDatabase("H:\db\db_SPEC\H_H_Jspec05_CURRENT.mdb") '12/22/17 'X_JSPEC_17_01
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_SPEC\X_JSPEC_17_01.mdb")
'**********************************************************************************************************
'**********************************************************************************************************

strPCSI = "SELECT PROJCSI.PROJCSIID, PROJCSI.PROJCSIDATE, PROJCSI.PROJSPECID, " _
& "PROJCSI.PROJ_NO, MCSI.MCSIID, MCSI.MCSI, MCSI.MCSINAME " _
& "FROM MCSI INNER JOIN PROJCSI ON MCSI.MCSIID = PROJCSI.MCSIID " _
& "WHERE (((PROJCSI.PROJCSIID)=" & xProjCSIID & "));"
'& "ORDER BY DRGP.DRGNO;"

Set RsCSI = DB.OpenRecordset(strPCSI)
rnCSI = RsCSI.RecordCount



With RsCSI
For iCSI = 1 To rnCSI
    Debug.Print "xprojcsiid : " & xProjCSIID
    UserForm1.TextBox1 = RsCSI!mCSI
    'strPpart = "SELECT PROJPART.PROJPARTID, PROJPART.PROJCSIID, PROJPART.PROJPARTNO, " _
    & "MPART.MPARTID, MPART.MPART FROM MPART INNER JOIN PROJPART ON PROJPART.MPARTID=mpart.MPARTID " _
    & "WHERE (((PROJPART.PROJCSIID)=" & xProjCSIID & "));"
    strPpart = "SELECT PROJPART.PROJPARTID, PROJPART.PROJCSIID, PROJPART.PROJPARTNO, MPART.MPARTID, MPART.MPART " _
    & "FROM MPART INNER JOIN PROJPART ON MPART.MPARTID = PROJPART.MPARTID " _
    & "WHERE (((PROJPART.PROJCSIID)=" & xProjCSIID & "))" _
    & "ORDER BY PROJPART.PROJPARTNO;"
    Set RsPart = DB.OpenRecordset(strPpart, dbOpenDynaset)
        With RsPart
        .MoveFirst
        .MoveLast
        .MoveFirst
        rnPart = RsPart.RecordCount
        Debug.Print "rnPart: " & rnPart
    
        For iPart = 1 To rnPart
            UserForm1.TextBox1 = UserForm1.TextBox1 & vbCrLf & RsPart!MPART
            xProjpartid = !projpartid
            strPat = "SELECT PROJAT.PROJATID, PROJAT.PROJPARTID, PROJAT.PROJATNO, " _
            & "MAT.MATID, MAT.MAT FROM MAT INNER JOIN PROJAT ON MAT.MATID=PROJAT.MATID " _
            & "WHERE (((PROJat.PROJpartid)=" & xProjpartid & "))" _
            & "ORDER BY PROJAT.PROJATNO;"
            Set Rsat = DB.OpenRecordset(strPat, dbOpenDynaset)
                With Rsat
                '.MoveFirst
                '.MoveLast
                '.MoveFirst
                rnAt = Rsat.RecordCount
                Debug.Print "rnat: " & rnAt
            
                For iat = 1 To rnAt
                    UserForm1.TextBox1 = UserForm1.TextBox1 & vbCrLf & Chr(9) & Rsat!Mat
                    xProjatid = !projatid
                    
                    strPpara1 = "SELECT PROJPARA1.PROJPARA1ID, PROJPARA1.PROJATID, PROJPARA1.PROJPARA1NO, " _
                    & "MPARA1.MPARA1ID, MPARA1.MPARA1 FROM PROJPARA1 INNER JOIN MPARA1 ON PROJPARA1.MPARA1ID=MPARA1.MPARA1ID " _
                    & "WHERE (((PROJpara1.PROJatid)=" & xProjatid & "));" _
                    & "ORDER BY PROJPARA1.PROJPARA1NO;"
                
                    Set Rspara1 = DB.OpenRecordset(strPpara1)
                        With Rspara1
                        .MoveLast
                        .MoveFirst
                            rnPara1 = Rspara1.RecordCount
                            For ipara1 = 1 To rnPara1
                                UserForm1.TextBox1 = UserForm1.TextBox1 & vbCrLf & Chr(9) & Chr(9) & Rspara1!Mpara1
                                xProjpara1id = !projpara1id
                                
                        '-------------------------------------------------------------------------
                                
                                Rspara1.MoveNext
                            Next ipara1
                        End With 'rspara1
                    
                    
                    '-------------------------------------------------------------------------
                    
                    Rsat.MoveNext
                    Next iat
                End With 'rsat
            
            
            '-------------------------------------------------------------------------
            RsPart.MoveNext
            Next iPart
        End With 'rspart
Next iCSI

End With 'rscsi



Debug.Print rnPart & rnAt & rnPara1
'Do Until RsPart.EOF = True

'    UserForm1.TextBox1 = RsPart!MPART
   ' RsPart.MoveNext
    
'Loop
End Function
'**********************************************************************************************************


'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'
'12/22/17
'
'***/*******************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************


Function writemtext(xProjCSIID As Integer)
'On Error GoTo Err_writemtext
Dim DB As Database
Dim RsPart, RsPat, RsPpara1 As Recordset
Dim strPpart, strPat, strPpara1, text As String
Dim rnPart, rnAt, rnPara1 As Integer
Dim specArray() As Variant
Dim corner(0 To 2) As Double
Dim width As Double
Dim mATno As Integer


'**********************************************************************************************************
'**********************************************************************************************************
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_SPEC\X_JSPEC_17_01.mdb")
'**********************************************************************************************************
'**********************************************************************************************************

strPCSI = "SELECT PROJCSI.PROJCSIID, PROJCSI.PROJCSIDATE, PROJCSI.PROJSPECID, " _
& "PROJCSI.PROJ_NO, MCSI.MCSIID, MCSI.MCSI, MCSI.MCSINAME " _
& "FROM MCSI INNER JOIN PROJCSI ON MCSI.MCSIID = PROJCSI.MCSIID " _
& "WHERE (((PROJCSI.PROJCSIID)=" & xProjCSIID & "));"


Set RsCSI = DB.OpenRecordset(strPCSI)
rnCSI = RsCSI.RecordCount

corner(0) = 0#: corner(1) = 0#: corner(2) = 0#
'WIDTH = 384 '32' X 12
'*********************************************************************************************
'*********************************************************************************************
' 12/22/17
'*********************************************************************************************
'*********************************************************************************************
If frmSpec.OptionButton1.value Then
    mSCF = 48
    width = mSCF * 8 '48 x 8 = 384
ElseIf frmSpec.OptionButton2.value Then
    mSCF = 96
    width = mSCF * 8 '96 x 8 = 768
ElseIf frmSpec.OptionButton3.value Then
    mSCF = 64
    width = mSCF * 8 '64 x 8 = 512
End If

xScF8 = mSCF / 12 * 2
xScF6 = mSCF / 12 * 1.5
xScF4 = mSCF / 12 * 1
'*********************************************************************************************
'*********************************************************************************************

text = ""
mATno = 1

With RsCSI
For iCSI = 1 To rnCSI
    Debug.Print "xprojcsiid : " & xProjCSIID
'*********************** Format Start CSIName ************************************************
    If frmSpec!mNote = True Then
        '*********************************************************************************************************
        '*********************************************************************************************************
        'text = text & Chr(10) & "\Farial.ttf;\H8.0;" & RsCSI!mCSI & " " & RsCSI!MCSIname & "\P"
        '*********************************************************************************************************
        text = text & Chr(10) & "\Farial.ttf;\H" & xScF8 & ";" & RsCSI!mCSI & " " & RsCSI!MCSIname & "\P"
        '*********************************************************************************************************
        '*********************************************************************************************************
    Else
        text = text & Chr(10) & RsCSI!mCSI & " " & RsCSI!MCSIname & "\P"
    End If
'*********************** Format End CSIName ************************************************

    'Set MTextObj = ThisDrawing.ModelSpace.AddMText(corner, width, text)

    'UserForm1.TextBox1 = RsCSI!MCSI
    
    'strPpart = "SELECT PROJPART.PROJPARTID, PROJPART.PROJCSIID, PROJPART.PROJPARTNO, " _
    & "MPART.MPARTID, MPART.MPART FROM MPART INNER JOIN PROJPART ON PROJPART.MPARTID=mpart.MPARTID " _
    & "WHERE (((PROJPART.PROJCSIID)=" & xProjCSIID & "));"
    strPpart = "SELECT PROJPART.PROJPARTID, PROJPART.PROJCSIID, PROJPART.PROJPARTNO, MPART.MPARTID, MPART.MPART " _
    & "FROM MPART INNER JOIN PROJPART ON MPART.MPARTID = PROJPART.MPARTID " _
    & "WHERE (((PROJPART.PROJCSIID)=" & xProjCSIID & ")) " _
    & "ORDER BY PROJPART.PROJPARTNO;"

    Set RsPart = DB.OpenRecordset(strPpart)
        With RsPart
        .MoveFirst
        .MoveLast
        .MoveFirst

        rnPart = RsPart.RecordCount
        If rnPart > 0 Then
        
        Debug.Print "rnPart: " & rnPart
        
        'text = text & vbCrLf & RsPart!MPART
        'Set MTextObj = ThisDrawing.ModelSpace.AddMText(corner, width, text)

        'For iPart = 1 To rnPart
        For iPart = 1 To rnPart
        
            'UserForm1.TextBox1 = UserForm1.TextBox1 & vbCrLf & RsPart!MPART
            ' Creates the mtext Object
            'text = text & vbCrLf & RsPart!MPART
'*********************** Format Start Part ************************************************
            If frmSpec!mNote = True Then
                
                'text = text & "\P" & Chr(9) & RsPart!MPART & "\P"
                Debug.Print RsPart!MPART
                text = text
            Else
                text = text & "\P" & "PART " & RsPart!projpartno & Chr(9) & RsPart!MPART & "\P"
                'Set MTextObj = ThisDrawing.ModelSpace.AddMText(corner, width, text)
            End If
'*********************** Format End Part ************************************************

            xProjpartid = !projpartid
            strPat = "SELECT PROJAT.PROJATID, PROJAT.PROJPARTID, PROJAT.PROJATNO, " _
            & "MAT.MATID, MAT.MAT FROM MAT INNER JOIN PROJAT ON MAT.MATID=PROJAT.MATID " _
            & "WHERE (((PROJat.PROJpartid)=" & xProjpartid & ")) " _
            & "ORDER BY PROJAT.PROJATNO;"

            Set Rsat = DB.OpenRecordset(strPat, dbOpenDynaset)
                With Rsat
                If Rsat.RecordCount = 0 Then
                    GoTo ExitRsAt
                End If
                .MoveFirst
                .MoveLast
                .MoveFirst

                rnAt = Rsat.RecordCount
                If rnAt > 0 Then
                
                Debug.Print "rnat: " & rnAt
            
                For iat = 1 To rnAt
'*********************** Format Start AT ************************************************
            If frmSpec!mNote = True Then
                'text = text & "\P" & Rsat!projatno & Chr(46) & Chr(9) & UCase(Rsat!Mat) & "\P"
                '*********************************************************************************************************
                '*********************************************************************************************************
                'text = text & "\P" & "\Farial.ttf;\H6.0;" & mATno & Chr(46) & Chr(32) & Chr(32) & UCase(Rsat!Mat) '& "\P"
                '*********************************************************************************************************
                text = text & "\P" & "\Farial.ttf;\H" & xScF6 & ";" & mATno & Chr(46) & Chr(32) & Chr(32) & UCase(Rsat!Mat) '& "\P"
                '*********************************************************************************************************
                '*********************************************************************************************************
                mATno = mATno + 1
                
            Else
                text = text & "\P" & Rsat!projatno & Chr(46) & Chr(9) & UCase(Rsat!Mat) & "\P"
                'Set MTextObj = ThisDrawing.ModelSpace.AddMText(corner, width, text)
            End If
'*********************** Format End AT ************************************************

                    'UserForm1.TextBox1 = UserForm1.TextBox1 & vbCrLf & Chr(9) & RsAt!MAT
                    xProjatid = !projatid
                    
                    strPpara1 = "SELECT PROJPARA1.PROJPARA1ID, PROJPARA1.PROJATID, PROJPARA1.PROJPARA1NO, " _
                    & "MPARA1.MPARA1ID, MPARA1.MPARA1 FROM PROJPARA1 INNER JOIN MPARA1 ON PROJPARA1.MPARA1ID=MPARA1.MPARA1ID " _
                    & "WHERE (((PROJpara1.PROJatid)=" & xProjatid & ")) " _
                    & "ORDER BY PROJPARA1.PROJPARA1NO;"
                
                    Set Rspara1 = DB.OpenRecordset(strPpara1, dbOpenDynaset)
                        With Rspara1
                        If Rspara1.RecordCount = 0 Then
                            GoTo ExitPara1
                        End If
                        .MoveFirst
                        .MoveLast
                        .MoveFirst

                        rnPara1 = Rspara1.RecordCount
                        If rnPara1 > 0 Then
                            
                            For ipara1 = 1 To rnPara1
'*********************** Format Start P1 ************************************************
                                If frmSpec!mNote = True Then
                                    'SendKeys "%(09)"
                                    '*********************************************************************************************************
                                    '*********************************************************************************************************
                                    'text = text & "\P" & "\Fromans.shx;\H3.5;" & Chr(9) & Chr(45) & Chr(9) & UCase(Rspara1!Mpara1) & "\P"
                                    '*********************************************************************************************************
                                    text = text & "\P" & "\Fromans.shx;\H" & xScF4 & ";" & Chr(9) & Chr(45) & Chr(9) & UCase(Rspara1!Mpara1) & "\P"
                                    '*********************************************************************************************************
                                    '*********************************************************************************************************
                                    '*********************************************************************************************************
                                
                                Else
                                    text = text & "\P" & Chr(9) & Rspara1!projpara1no & Chr(46) & Chr(9) & UCase(Rspara1!Mpara1) & "\P"
                                'Set MTextObj = ThisDrawing.ModelSpace.AddMText(corner, width, text)
                                End If
'*********************** Format End P1 ************************************************
                                'UserForm1.TextBox1 = UserForm1.TextBox1 & vbCrLf & Chr(9) & Chr(9) & rsPara1!Mpara1
                                xProjpara1id = !projpara1id
                                
                                strPpara2 = "SELECT PROJPARA2.PROJPARA2ID, PROJPARA2.PROJPARA1ID, PROJPARA2.PROJPARA2NO,  MPARA2.MPARA2ID, " _
                                & "MPARA2.MPARA2 FROM PROJPARA2 INNER JOIN MPARA2 ON PROJPARA2.MPARA2ID=MPARA2.MPARA2ID " _
                                & "WHERE (((PROJpara2.PROJpara1id)=" & xProjpara1id & ")) " _
                                & "ORDER BY PROJPARA2.PROJPARA2id;"
                                '& "ORDER BY PROJPARA2.PROJPARA2id;"
                                '& "ORDER BY PROJPARA2.PROJPARA2NO;"
                                
                                'PROJPARA2.IIf([PROJPARA2NO] Is Null,0,Val([projpara2no])) AS PROJPARA2NOv,
                                
                                
                                Set rspara2 = DB.OpenRecordset(strPpara2, dbOpenDynaset)
                                With rspara2
                                If rspara2.RecordCount = 0 Then
                                    GoTo ExitPara2
                                End If
                                
                                rnpara2 = rspara2.RecordCount
                                '.Sort = "val[PROJPARA2.PROJPARA2no] asC"
                                If rnpara2 > 0 Then
                                .MoveFirst
                                .MoveLast
                                .MoveFirst
                                .Sort = "val[PROJPARA2.PROJPARA2no] asC"
                                    rnpara2 = rspara2.RecordCount
                                    For ipara2 = 1 To rnpara2
'*********************** Format Start P2 ************************************************
                                    If frmSpec!mNote = True Then
                                        text = text & "\P" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(45) & Chr(9) & UCase(rspara2!Mpara2) & "\P"
                                    Else
                                        text = text & "\P" & Chr(9) & Chr(9) & rspara2!PROJPARA2no & Chr(46) & Chr(9) & UCase(rspara2!Mpara2) & "\P"
                                        Debug.Print rspara2!PROJPARA2no
                                    End If
'*********************** Format End P2 ************************************************
                                        xProjpara2id = !projpara2id
                                '-------------------------------------------------------------------------
                                        strPpara3 = "SELECT PROJPARA3.PROJPARA3ID, PROJPARA3.PROJPARA2ID, PROJPARA3.PROJPARA3NO, MPARA3.MPARA3ID, " _
                                        & "MPARA3.MPARA3 FROM PROJPARA3 INNER JOIN MPARA3 ON PROJPARA3.MPARA3ID=MPARA3.MPARA3ID " _
                                        & "WHERE (((PROJpara3.PROJpara2id)=" & xProjpara2id & ")) " _
                                        & "ORDER BY PROJPARA3.PROJPARA3NO;"
                                        Set rspara3 = DB.OpenRecordset(strPpara3, dbOpenDynaset)
                                        
                                        
                                        With rspara3
                                        If rspara3.RecordCount = 0 Then
                                        GoTo ExitPara3
                                        End If
                                        .MoveFirst
                                        .MoveLast
                                        .MoveFirst

                                        rnpara3 = rspara3.RecordCount
                                        Debug.Print rnpara3
                                        If rnpara3 > 0 Then
                                            
                                            For ipara3 = 1 To rnpara3
'*********************** Format Start P3 ************************************************
                                                If frmSpec!mNote = True Then
                                                    text = text & "\P" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(45) & Chr(9) & rspara3!Mpara3 & "\P"
                                                Else
                                                    text = text & "\P" & Chr(9) & Chr(9) & Chr(9) & rspara3!projpara3no & Chr(46) & Chr(9) & rspara3!Mpara3 & "\P"
                                                End If
'*********************** Format End P3 ************************************************
                                                Debug.Print ">>>projpara3id>>> "; !projpara3id
                                                xProjpara3id = !projpara3id
                                                If !projpara3id = 2147 Then
                                                Debug.Print ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> check"
                                                End If
                                '-------------------------------------------------------------------------
                                '-------------------------------------------------------------------------
                                                strPpara4 = "SELECT PROJPARA4.PROJPARA4ID, PROJPARA4.PROJPARA3ID, PROJPARA4.PROJPARA4NO, MPARA4.MPARA4ID, " _
                                                & "MPARA4.MPARA4 FROM PROJPARA4 INNER JOIN MPARA4 ON PROJPARA4.MPARA4ID=MPARA4.MPARA4ID " _
                                                & "WHERE (((PROJpara4.PROJpara3id)=" & xProjpara3id & ")) " _
                                                & "ORDER BY PROJPARA4.PROJPARA4NO;"
                                                Set rspara4 = DB.OpenRecordset(strPpara4, dbOpenDynaset)
                                
                                                With rspara4
                                                    If rspara4.RecordCount = 0 Then
                                                        GoTo ExitPara4
                                                    End If
                                                .MoveFirst
                                                .MoveLast
                                                .MoveFirst
        
                                                rnpara4 = rspara4.RecordCount
                                                If rnpara4 > 0 Then
                                                    
                                                    For ipara4 = 1 To rnpara4
                                                        text = text & "\P" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & rspara4!projpara4no & Chr(46) & Chr(9) & rspara4!Mpara4 & "\P"
                                                        xProjpara4id = !projpara4id
                                                Debug.Print text
                                                        rspara4.MoveNext
                                                    Next ipara4
                                                End If
ExitPara4:
                                                End With 'rspara4
                                                rspara4.Close
                                '-------------------------------------------------------------------------
                                '-------------------------------------------------------------------------
                                                rspara3.MoveNext
                                            Next ipara3
                                        End If
ExitPara3:
                                        End With 'rspara3
                                        rspara3.Close
                        '-------------------------------------------------------------------------
                                        rspara2.MoveNext
                                    Next ipara2
                                    End If
                                    
ExitPara2:
                                End With 'rspara2
                                
                                rspara2.Close
                                
                                
                                
                        '-------------------------------------------------------------------------
                                
                                Rspara1.MoveNext
                            Next ipara1
                            End If
ExitPara1:
                        End With 'rspara1
                        Rspara1.Close
                    
                    '-------------------------------------------------------------------------
                    
                    Rsat.MoveNext
                    Next iat
                End If
ExitRsAt:
                End With 'rsat
                Rsat.Close
            
            '-------------------------------------------------------------------------
            RsPart.MoveNext
            Next iPart
            End If
ExitRsPart:
        End With 'rspart
        RsPart.Close
Next iCSI
Set mtextObj = ThisDrawing.ModelSpace.AddMText(corner, width, text)
If frmSpec!mNote = True Then
    CreateARstyle
    mtextObj.StyleName = "ar"
    mtextObj.height = 4
Else
    CreateCNstyle
    mtextObj.StyleName = "cn"
    mtextObj.height = 4
End If
mtextObj.Update
RsCSI.Close
End With 'rscsi

DB.Close


Debug.Print rnPart & rnAt & rnPara1
'Do Until RsPart.EOF = True

'    UserForm1.TextBox1 = RsPart!MPART
   ' RsPart.MoveNext
    
'Loop
'Exit_writemtext:
    'Exit Function

'Err_writemtext:
    'MsgBox Err.Description
    'Resume Exit_writemtext
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



