Attribute VB_Name = "Mod_Sched"
Option Explicit
Public Sub SCHEDULES2()
   frmSched_CSI.show 'MOD 9/9/14
End Sub
Public Sub SCHEDULES()
   frmDoor.show
End Sub

Sub test()
Dim x As String
x = SEPTEXT2("UN CERAMIC MOSAIC tile", 2, 2, 12)
Debug.Print "x: " & x
End Sub

Sub mkDr()


Dim DB As Database
Dim RsDRG As Recordset
Dim strDRG As String
Dim rnDRG As Integer
Dim rcArray As Variant
Dim TblPT(0 To 2) As Double
Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
Dim height As Double
Dim mode As Long
Dim prompt, txtRow As String
Dim tag As String
Dim value As String
Dim blockRefObj As AcadBlockReference
Dim attData() As AcadObject
Dim nBubdwg As Variant
Dim EntGrp(0), i, ati As Integer
Dim EntPrp(0) As Variant
Dim BlkObj As Object
Dim Pt1(0) As Double
Dim Pt2(0) As Double
'ThisDrawing.SelectionSets.Item("TBLK").Delete
'create a selection set
'Set SSNew = ThisDrawing.SelectionSets.Add("TBLK")
'Filter for Group code 2, the block name
'EntGrp(0) = 2
'The name of the block to filter for
'EntPrp(0) = "xdrschedrow01"
'find the block
'SSNew.Select acSelectionSetAll, Pt1, Pt2, EntGrp, EntPrp
'If a block is found
'If SSNew.Count >= 1 Then






frmDoor.ListBox1.Clear
CreateCNstyle


Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\Door02_01.mdb")

strDRG = "SELECT DRGP.DRGID, DRGP.DRID, DRGP.PROJ_NO, DRGP.DRGNO, DRGP.DRTYPE, " _
& "DRGP.DRFIN, DRGP.DRFRAME, DRGP.DRJAMB, DRGP.DRHEAD, DRGP.HDSET, DRGP.DRGNOTE, " _
& "DR.DRWIDTH, DR.DRHEIGHT, DR.DRMAT, DR.DRTHK, DR.DRMFGR, DR.DRFIRE, DR.DRNOTE, DR.OLDSIZE " _
& "FROM DRGP INNER JOIN DR ON DRGP.DRID = DR.DRID " _
& "WHERE (((DRGP.PROJ_NO) = '" & frmDoor.seleproj & "')) " _
& "ORDER BY DRGP.DRGNO;"

'drgno drtype,
'nBubdwg = "H:\drawfile\jvba\XBUBdre_01.dwg"  'MOD 5/18/15

'**************************************************************
'nBubdwg = "H:\drawfile\jvba\XBUBdrAR.dwg"  'MOD 6/5/15
'InsertJBlock "e:\drawfile\jvba\xdrschedrow02.dwg", tblPT, "Model"
'**************************************************************
'**************************************************************
nBubdwg = "\\jwja-svr-10\drawfile\jvba\XBUBdrAR.dwg"  'MOD 6/5/15
InsertJBlock "\\jwja-svr-10\drawfile\jvba\xdrschedrow02.dwg", TblPT, "Model"
'**************************************************************




Set RsDRG = DB.OpenRecordset(strDRG)
If RsDRG.RecordCount = 0 Then
GoTo ExitMKDR
End If
rnDRG = RsDRG.RecordCount
Debug.Print rnDRG
TblPT(0) = 0: TblPT(1) = 0: TblPT(1) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify
'***********************************************************************************************
'ThisDrawing.ModelSpace.InsertBlock tblPT, "e:\drawfile\jvba\xdrschedhead01.dwg", 1, 1, 1, 0
'ThisDrawing.ModelSpace.InsertBlock tblPT, "H:\drawfile\jvba\xdrschedhead02.dwg", 1, 1, 1, 0
ThisDrawing.ModelSpace.InsertBlock TblPT, "\\jwja-svr-10\drawfile\jvba\xdrschedhead02.dwg", 1, 1, 1, 0
'***********************************************************************************************
TblPT(1) = TblPT(1) - 144

RsDRG.MoveFirst
With RsDRG
    For i = 1 To rnDRG
         Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, "xdrschedrow02", 1, 1, 1, 0)

        'Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
               (insertionPnt, "xdrschedrow01", 1#, 1#, 1#, 0)
        'set blcokrefobj = thisdrawing.modelspace.Name.ObjectName "xdrschedrow01"
         
        '******************************
          'Get the block's attributes
                Dim varAttributes As Variant
                varAttributes = blockRefObj.GetAttributes
               
                    For ati = LBound(varAttributes) To UBound(varAttributes)
                    Debug.Print varAttributes(ati).TagString
                        If varAttributes(ati).TagString = "XDrno" Then
                            If IsNull(!DRGNO) = True Then
                                varAttributes(ati).textString = ""
                            Else
                                varAttributes(ati).textString = !DRGNO
                                '***************************************************************
                                '*      INSERT BUBBLE
                                '***************************************************************
                                                TblPT(0) = TblPT(0) - 48
                                                Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                             (TblPT, nBubdwg, 1, 1, 1, 0)
                                                Dim varAttributesB As Variant
                                                varAttributesB = blockRefObj.GetAttributes
                                                varAttributesB(0).textString = ![DRGNO]
                                                
                                                TblPT(0) = TblPT(0) + 48
                                
                                '***************************************************************
                            End If

                        End If
                        If varAttributes(ati).TagString = "xdrsize" Then
                            If IsNull(!OLDSIZE) = True Then
                                varAttributes(ati).textString = ""
                            Else
                                varAttributes(ati).textString = !OLDSIZE
                            End If

                        End If
                        If varAttributes(ati).TagString = "XDRTYPE" Then
                            If IsNull(!DRTYPE) = True Then
                                varAttributes(ati).textString = ""
                            Else
                                varAttributes(ati).textString = !DRTYPE
                            End If

                        End If
                        If varAttributes(ati).TagString = "XDRTHK" Then
                            If IsNull(!DRTHK) = True Then
                                varAttributes(ati).textString = ""
                            Else
                                varAttributes(ati).textString = !DRTHK
                            End If

                        End If
                        If varAttributes(ati).TagString = "XDRCON" Then
                            If IsNull(!DRMAT) = True Then
                                varAttributes(ati).textString = ""
                            Else
                                varAttributes(ati).textString = !DRMAT
                            End If
                        End If
                        If varAttributes(ati).TagString = "XDRFAC" Then
                            If IsNull(!DRFIN) = True Then
                                varAttributes(ati).textString = ""
                            Else
                                varAttributes(ati).textString = !DRFIN
                            End If
                        End If
        '***************************************************
                        If varAttributes(ati).TagString = "XFRTYPE" Then
                            If IsNull(!DRFRAME) = True Then
                                varAttributes(ati).textString = ""
                            Else
                                varAttributes(ati).textString = !DRFRAME
                            End If
                        End If
        '***************************************************
                        If varAttributes(ati).TagString = "XDRJ" Then
                            If IsNull(!drjamb) = True Then
                                varAttributes(ati).textString = ""
                            Else
                                varAttributes(ati).textString = !drjamb
                            End If
                        End If
                        If varAttributes(ati).TagString = "XDRH" Then
                            If IsNull(!DRHEAD) = True Then
                                varAttributes(ati).textString = ""
                            Else
                                varAttributes(ati).textString = !DRHEAD
                            End If
                        End If
                        If varAttributes(ati).TagString = "XDRF" Then
                            If IsNull(!DRFIRE) = True Then
                                varAttributes(ati).textString = ""
                            Else
                                varAttributes(ati).textString = !DRFIRE
                            End If
                        End If
                        If varAttributes(ati).TagString = "XDRHD" Then
                            If IsNull(!HDSET) = True Then
                                varAttributes(ati).textString = ""
                            Else
                                varAttributes(ati).textString = !HDSET
                            End If
                        End If
                        If varAttributes(ati).TagString = "XDRREM" Then
                            If IsNull(!DRGNOTE) = True Then
                                varAttributes(ati).textString = ""
                            Else
                                varAttributes(ati).textString = !DRGNOTE
                            End If
                        End If

                    Next ati
                'attributeObj.Update
        
                '************************************************
                'txtRow = !drgno & Chr(9) & !oldsize
                'Set textObj = ThisDrawing.ModelSpace.AddText(txtRow, tblPT, txtHT)
                '******************************************************
                'textObj.StyleName = "cn"
                'textObj.Update
                TblPT(1) = TblPT(1) - 12
        .MoveNext
        
    Next i
End With
'ThisDrawing.SelectionSets.Item("TBLK").Delete
RsDRG.Close
Set DB = Nothing
'Else
'ThisDrawing.SelectionSets.Item("TBLK").Delete
'no attribute block, inform the user

'MsgBox "Inserted block! Re-run this program.", vbCritical, "JWJA"

'delete the selection set
'ThisDrawing.SelectionSets.Item("TBLK").Delete
'End If

'Set MTextObj = ThisDrawing.ModelSpace.AddMText(corner, width, text)

'ThisDrawing.SelectionSets.Item("TBLK").Delete


ExitMKDR:
frmDoor.Hide
End Sub
Public Sub MkHDWR()


Dim DB As Database
Dim RstHDProjSet, RsHDwrSet As Recordset
Dim strHDSet, strHDProjSet, txtRow As String
Dim rnHDPS As Integer   'ProjSet no.
Dim rnHDS As Integer    'EachSet No.
'Dim rcArray As Variant
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
Dim height As Double
Dim mode As Long
Dim prompt As String
'Dim tag As String
'Dim value As String
Dim blockRefObj As AcadBlockReference
'Dim attData() As AcadObject

'Dim FilterGp(0) As Integer
'Dim FilterDt(0) As Variant
Dim BlkObj As Object
'Dim Pt1(0) As Double
'Dim Pt2(0) As Double

Dim xHDWRSETID, i, iatHG, rnHDWR, k, iatHW As Integer
'*************************************************************************
'*************************************************************************
Dim nHeadDwg, nGroupDwg, nRowDwg, nFootDwg As Variant
Dim nJumpHead, nJumpGroup, nJumpRow As Integer

'*************************************************************************
'*************************************************************************


    'If Not IsNull(ThisDrawing.SelectionSets.Item("hdwrsched")) Then
        'Set SSNew = ThisDrawing.SelectionSets.Item("hdwrsched")
        'SSNew.Delete
    'End If

'create a selection set
'ThisDrawing.SelectionSets.Item("hdwrsched").Delete

''Set SSNew = ThisDrawing.SelectionSets.Add("hdwrsched")
'Filter Type: for Group code 2, the block name
''FilterGp(0) = 2
'Filter Data: The name of the block to filter for
''FilterDt(0) = "XHDWRROW01"
'find the block
'SSNew.Select acSelectionSetAll, Pt1, Pt2, FilterGp, FilterDt
'If a block is found
''If SSNew.Count >= 1 Then

'If SSNew.Count >= 0 Then




frmDoor.ListBox1.Clear
CreateCNstyle


Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\Door02_01.mdb")

strHDProjSet = "SELECT HDWRSET.HDWRSETID, HDWRSET.PROJ_NO, HDWRSET.HDWRSETNOTE, HDWRSET.HDWRSETNO " _
& "FROM HDWRSET " _
& "WHERE (((HDWRSET.PROJ_NO)='" & frmDoor.seleproj & "'));"

'strHDSet = "SELECT HDWRSET.HDWRSETID, HDWRCONN.SETQTY, HDWR.TYPE, HDWR.CAT, HDWR.MFGR, HDWR.FIN, HDWR.HDWRNOTE, HDWR.HDIMGNAME, HDWR.HDWRDESC " _
& "FROM HDWRSET INNER JOIN (HDWR INNER JOIN HDWRCONN ON HDWR.HDID = HDWRCONN.HDID) ON HDWRSET.HDWRSETID = HDWRCONN.HDWRSETID " _
& "WHERE (((HDWRSET.HDWRSETID)=" & xHDWRSETID & "));"

strHDSet = "SELECT HDWRCONN.HDWRSETID, " _
                        & "HDWRCONN.SETQTY, HDWR.TYPE, HDWR.CAT, HDWR.MFGR, HDWR.FIN, HDWR.HDWRNOTE, HDWR.HDIMGNAME, HDWR.HDWRDESC " _
                        & "FROM HDWR INNER JOIN HDWRCONN ON HDWR.HDID = HDWRCONN.HDID " _
                        & "WHERE (((HDWRCONN.HDWRSETID)=" & xHDWRSETID & "));"

'& "WHERE (((HDWRSET.HDWRSETID)=576));"
'" _
'& "WHERE (((HDWRSET.HDWRSETID)=" & xHDWRSETID & "))"
'& "WHERE (((HDWRSET.HDWRSETID)=578));"


RecordSets:
Set RstHDProjSet = DB.OpenRecordset(strHDProjSet, dbOpenDynaset)

If RstHDProjSet.RecordCount = 0 Then
    GoTo ExitMKHDWR
End If


TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify
'InsertJBlock "e:\drawfile\jvba\XHDWRHEAD01.dwg", tblPT, "Model"

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XHDWRHEAD02.dwg"
    nGroupDwg = "\\jwja-svr-10\drawfile\jvba\XHDWRGPROW02.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\XHDWRROW02.dwg"
    nFootDwg = "\\jwja-svr-10\drawfile\jvba\XHDWRFOOT02.dwg"
    nJumpHead = 24
    nJumpGroup = 36
    nJumpRow = 12
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\XHDWRHEAD03.dwg", TblPT, "Model"

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


ThisDrawing.ModelSpace.InsertBlock TblPT, "\\jwja-svr-10\drawfile\jvba\XHDWRHEAD02.dwg", 1, 1, 1, 0
InsertJBlock "\\jwja-svr-10\drawfile\jvba\XHDWRGPROW02.dwg", tblPT2, "Model"
InsertJBlock "\\jwja-svr-10\drawfile\jvba\XHDWRROW02.dwg", tblPT2, "Model"
InsertJBlock "\\jwja-svr-10\drawfile\jvba\XHDWRFOOT02.dwg", tblPT2, "Model"

'ThisDrawing.SelectionSets.Item("XHDWRGPROW01").Delete
'ThisDrawing.SelectionSets.Item("XHDWRROW01").Delete

TblPT(1) = TblPT(1) - 24
RstHDProjSet.MoveLast
RstHDProjSet.MoveFirst
rnHDPS = RstHDProjSet.RecordCount
Debug.Print rnHDPS

With RstHDProjSet
    For i = 1 To rnHDPS
         xHDWRSETID = !HDWRSETID
         Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, "XHDWRGPROW02", 1, 1, 1, 0)

        'Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
               (insertionPnt, "xdrschedrow01", 1#, 1#, 1#, 0)
        'set blcokrefobj = thisdrawing.modelspace.Name.ObjectName "xdrschedrow01"
         
        '******************************
          'Get the block's attributes
                Dim varAttributes As Variant
                varAttributes = blockRefObj.GetAttributes
               
                    For iatHG = LBound(varAttributes) To UBound(varAttributes)
                    
                    Debug.Print varAttributes(iatHG).TagString
                        If varAttributes(iatHG).TagString = "XHDSETNO" Then
                            If IsNull(!HDWRSETNO) = True Then
                                varAttributes(iatHG).textString = ""
                            Else
                                varAttributes(iatHG).textString = !HDWRSETNO
                                Debug.Print varAttributes(iatHG).textString
                            End If

                        End If
                        If varAttributes(iatHG).TagString = "XHDSETNOTE" Then
                            If IsNull(!HDWRSETNOTE) = True Then
                                varAttributes(iatHG).textString = ""
                            Else
                                varAttributes(iatHG).textString = !HDWRSETNOTE
                            End If

                        End If
                    Next iatHG
                        TblPT(1) = TblPT(1) - 36
                        xHDWRSETID = !HDWRSETID
                        Debug.Print xHDWRSETID
                        'RstHDSet.Close
                        Set RsHDwrSet = DB.OpenRecordset("SELECT HDWRCONN.HDWRSETID, " _
                        & "HDWRCONN.SETQTY, HDWR.TYPE, HDWR.CAT, HDWR.MFGR, HDWR.FIN, HDWR.HDIMGNAME, HDWR.HDWRDESC " _
                        & "FROM HDWR RIGHT JOIN HDWRCONN ON HDWR.HDID = HDWRCONN.HDID " _
                        & "WHERE (((HDWRCONN.HDWRSETID)=" & xHDWRSETID & "));")
 '& xHDWRSETID & "));", dbOpenDynaset)
                        Debug.Print RsHDwrSet.RecordCount
                        'Debug.Print strHDSet

                        
                        If RsHDwrSet.RecordCount = 0 Then
                            GoTo ExitMKHDWR
                        End If
                        RsHDwrSet.MoveLast
                        RsHDwrSet.MoveFirst
                        Debug.Print RsHDwrSet.RecordCount
                        Debug.Print rnHDWR
                        rnHDWR = RsHDwrSet.RecordCount
                        With RsHDwrSet
                        Debug.Print xHDWRSETID
                            For k = 1 To rnHDWR
                                 Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                             (TblPT, "XHDWRROW02", 1, 1, 1, 0)

                                Dim varAttributes2 As Variant
                                varAttributes2 = blockRefObj.GetAttributes
                                For iatHW = LBound(varAttributes2) To UBound(varAttributes2)
                                Debug.Print varAttributes2(iatHW).TagString
                                    If varAttributes2(iatHW).TagString = "XQTY" Then
                                        If IsNull(!SETQTY) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            varAttributes2(iatHW).textString = !SETQTY
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XTYPE" Then
                                        If IsNull(!Type) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            varAttributes2(iatHW).textString = !Type
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XCAT" Then
                                        If IsNull(!cat) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            varAttributes2(iatHW).textString = !cat
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XFIN" Then
                                        If IsNull(!FIN) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            varAttributes2(iatHW).textString = !FIN
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XMFGR" Then
                                        If IsNull(!MFGR) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            varAttributes2(iatHW).textString = !MFGR
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XNOTE" Then
                                        If IsNull(!HDWRDESC) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            varAttributes2(iatHW).textString = Left(!HDWRDESC, 22)
                                        End If
                                    End If
                                Next iatHW
                            
                            TblPT(1) = TblPT(1) - 12
                            .MoveNext
                            Next k
                            End With 'hdwr rows
                   RsHDwrSet.Close
                'attributeObj.Update
        
                '************************************************
                'txtRow = !drgno & Chr(9) & !oldsize
                'Set textObj = ThisDrawing.ModelSpace.AddText(txtRow, tblPT, txtHT)
                '******************************************************
                'textObj.StyleName = "cn"
                'textObj.Update
                Debug.Print "iathw :" & iatHW
        TblPT(1) = TblPT(1)
        .MoveNext
        
    Next i
End With
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                             (TblPT, "XHDWRFOOT02", 1, 1, 1, 0)
''ThisDrawing.SelectionSets.Item("hdwrsched").Delete
RstHDProjSet.Close
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


ExitMKHDWR:
frmDoor.Hide

End Sub

'############################################################################
'############################################################################
'#
'#
'#          MkHDWRimg
'#
'#
'#
'############################################################################
'############################################################################

 Public Sub MkHDWRimg()


Dim DB As Database
Dim RstHDProjSet, RsHDwrSet As Recordset
Dim strHDSet, strHDProjSet, txtRow As String
Dim rnHDPS As Integer   'ProjSet no.
Dim rnHDS As Integer    'EachSet No.
'Dim rcArray As Variant
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
Dim height As Double
Dim mode As Long
Dim prompt As String

'**********IMG********************
Dim nHeadDwg, nGroupDwg, nRowDwg, nFootDwg As Variant
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

Dim xHDWRSETID, i, iatHG, rnHDWR, k, iatHW As Integer

    'If Not IsNull(ThisDrawing.SelectionSets.Item("hdwrsched")) Then
        'Set SSNew = ThisDrawing.SelectionSets.Item("hdwrsched")
        'SSNew.Delete
    'End If

'create a selection set
'ThisDrawing.SelectionSets.Item("hdwrsched").Delete

''Set SSNew = ThisDrawing.SelectionSets.Add("hdwrsched")
'Filter Type: for Group code 2, the block name
''FilterGp(0) = 2
'Filter Data: The name of the block to filter for
''FilterDt(0) = "XHDWRROW01"
'find the block
'SSNew.Select acSelectionSetAll, Pt1, Pt2, FilterGp, FilterDt
'If a block is found
''If SSNew.Count >= 1 Then

'If SSNew.Count >= 0 Then




frmDoor.ListBox1.Clear
CreateCNstyle


Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\Door02_01.mdb")

strHDProjSet = "SELECT HDWRSET.HDWRSETID, HDWRSET.PROJ_NO, HDWRSET.HDWRSETNOTE, HDWRSET.HDWRSETNO " _
& "FROM HDWRSET " _
& "WHERE (((HDWRSET.PROJ_NO)='" & frmDoor.seleproj & "'))" _
& "ORDER BY HDWRSET.HDWRSETNO ;"

'strHDSet = "SELECT HDWRSET.HDWRSETID, HDWRCONN.SETQTY, HDWR.TYPE, HDWR.CAT, HDWR.MFGR, HDWR.FIN, HDWR.HDWRNOTE, HDWR.HDIMGNAME, HDWR.HDWRDESC " _
& "FROM HDWRSET INNER JOIN (HDWR INNER JOIN HDWRCONN ON HDWR.HDID = HDWRCONN.HDID) ON HDWRSET.HDWRSETID = HDWRCONN.HDWRSETID " _
& "WHERE (((HDWRSET.HDWRSETID)=" & xHDWRSETID & "));"

strHDSet = "SELECT HDWRCONN.HDWRSETID, " _
                        & "HDWRCONN.SETQTY, HDWR.TYPE, HDWR.CAT, HDWR.MFGR, HDWR.FIN, HDWR.HDWRNOTE, HDWR.HDIMGNAME, HDWR.HDWRDESC " _
                        & "FROM HDWR INNER JOIN HDWRCONN ON HDWR.HDID = HDWRCONN.HDID " _
                        & "WHERE (((HDWRCONN.HDWRSETID)=" & xHDWRSETID & "));"

'& "WHERE (((HDWRSET.HDWRSETID)=576));"
'" _
'& "WHERE (((HDWRSET.HDWRSETID)=" & xHDWRSETID & "))"
'& "WHERE (((HDWRSET.HDWRSETID)=578));"


RecordSets:
Set RstHDProjSet = DB.OpenRecordset(strHDProjSet, dbOpenDynaset)

If RstHDProjSet.RecordCount = 0 Then
    GoTo ExitMKHDWRimg
End If


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

If frmDoor.Showhdwrimg = True Then
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XHDWRHEAD03.dwg"
    nGroupDwg = "\\jwja-svr-10\drawfile\jvba\XHDWRGPROW03.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\XHDWRROW03.dwg"
    nFootDwg = "\\jwja-svr-10\drawfile\jvba\XHDWRFOOT03.dwg"
    nJumpHead = 60 'was 24
    nJumpGroup = 24 'was 66
    nJumpRow = 60
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\XHDWRHEAD03.dwg", TblPT, "Model"
Else
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XHDWRHEAD02.dwg"
    nGroupDwg = "\\jwja-svr-10\drawfile\jvba\XHDWRGPROW02.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\XHDWRROW02.dwg"
    nFootDwg = "\\jwja-svr-10\drawfile\jvba\XHDWRFOOT02.dwg"
    nJumpHead = 24
    nJumpGroup = 36
    nJumpRow = 12
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\XHDWRHEAD02.dwg", TblPT, "Model"
End If

TblPT(1) = TblPT(1) - nJumpHead
RstHDProjSet.MoveLast
RstHDProjSet.MoveFirst
rnHDPS = RstHDProjSet.RecordCount
Debug.Print rnHDPS

With RstHDProjSet
    For i = 1 To rnHDPS
         xHDWRSETID = !HDWRSETID
         Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, nGroupDwg, 1, 1, 1, 0)

        'Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
               (insertionPnt, "xdrschedrow01", 1#, 1#, 1#, 0)
        'set blcokrefobj = thisdrawing.modelspace.Name.ObjectName "xdrschedrow01"
         
        '******************************
          'Get the block's attributes
                Dim varAttributes As Variant
                varAttributes = blockRefObj.GetAttributes
               
                    For iatHG = LBound(varAttributes) To UBound(varAttributes)
                    
                    Debug.Print varAttributes(iatHG).TagString
                        If varAttributes(iatHG).TagString = "XHDSETNO" Then
                            If IsNull(!HDWRSETNO) = True Then
                                varAttributes(iatHG).textString = ""
                            Else
                                varAttributes(iatHG).textString = !HDWRSETNO
                                Debug.Print varAttributes(iatHG).textString
                            End If

                        End If
                        If varAttributes(iatHG).TagString = "XHDSETNOTE" Then
                            If IsNull(!HDWRSETNOTE) = True Then
                                varAttributes(iatHG).textString = ""
                            Else
                                varAttributes(iatHG).textString = !HDWRSETNOTE
                            End If

                        End If
                    Next iatHG
                        TblPT(1) = TblPT(1) - nJumpGroup
                        xHDWRSETID = !HDWRSETID
                        Debug.Print xHDWRSETID
                        'RstHDSet.Close
                        Set RsHDwrSet = DB.OpenRecordset("SELECT HDWRCONN.HDWRSETID, " _
                        & "HDWRCONN.SETQTY, HDWR.TYPE, HDWR.CAT, HDWR.MFGR, HDWR.FIN, HDWR.HDIMGNAME, hdwr.hdwrnote, HDWR.HDWRDESC " _
                        & "FROM HDWR RIGHT JOIN HDWRCONN ON HDWR.HDID = HDWRCONN.HDID " _
                        & "WHERE (((HDWRCONN.HDWRSETID)=" & xHDWRSETID & "));")
 '& xHDWRSETID & "));", dbOpenDynaset)
                        Debug.Print RsHDwrSet.RecordCount
                        'Debug.Print strHDSet

                        
                        If RsHDwrSet.RecordCount = 0 Then
                            GoTo ExitMKHDWRimg
                        End If
                        RsHDwrSet.MoveLast
                        RsHDwrSet.MoveFirst
                        Debug.Print RsHDwrSet.RecordCount
                        Debug.Print rnHDWR
                        rnHDWR = RsHDwrSet.RecordCount
                        With RsHDwrSet
                        Debug.Print xHDWRSETID
                            For k = 1 To rnHDWR
                                 Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                             (TblPT, nRowDwg, 1, 1, 1, 0)

                                Dim varAttributes2 As Variant
                                varAttributes2 = blockRefObj.GetAttributes
                        '***************************************************************
                                If !hdimgname <> "" Then
                                If frmDoor.Showhdwrimg = True Then
                                    Dim imgPT(0 To 2) As Double
                                    Dim scalefactor, xSF As Double
                                    Dim rotAngleInDegree As Double, rotAngle As Double
                                    'Dim imageName As String
                                    Dim imageName As String
                                
                                    Dim raster As AcadRasterImage
                                    Dim imgheight As Variant
                                    Dim imgwidth As Variant
                                
                                    'imageName = "C:\AutoCAD\sample\downtown.jpg"
                                    imageName = !hdimgname
                                    Debug.Print imageName
                                    'insertionPoint(0) = 2#: insertionPoint(1) = 2#: insertionPoint(2) = 0#
                                    'imgPT(0) = tblPT(0) + 36: imgPT(1) = tblPT(1) - 54: imgPT(2) = tblPT(2)
                                    imgPT(0) = TblPT(0) + 6: imgPT(1) = TblPT(1) - 54: imgPT(2) = TblPT(2)
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

                                
                                
                                
                                For iatHW = LBound(varAttributes2) To UBound(varAttributes2)
                                Debug.Print varAttributes2(iatHW).TagString
                                    If varAttributes2(iatHW).TagString = "XQTY" Then
                                        If IsNull(!SETQTY) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            varAttributes2(iatHW).textString = !SETQTY
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XTYPE" Then
                                        If IsNull(!Type) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            varAttributes2(iatHW).textString = !Type
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XCAT" Then
                                        If IsNull(!cat) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            varAttributes2(iatHW).textString = !cat
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XFIN" Then
                                        If IsNull(!FIN) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            varAttributes2(iatHW).textString = !FIN
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XMFGR" Then
                                        If IsNull(!MFGR) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            varAttributes2(iatHW).textString = !MFGR
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XNOTE1" Then
                                        If IsNull(!hdwrnote) = True Then
                                        GoTo SkipforNull
                                        End If
                                        If IsNull(SEPTEXT2(!hdwrnote, 1, 2, 48)) = True Or !hdwrnote = "" Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            'varAttributes2(iatHW).TextString = Left(!HDWRDESC, 22)
                                            varAttributes2(iatHW).textString = SEPTEXT2(!hdwrnote, 1, 2, 48)
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XNOTE2" Then
                                        If IsNull(!hdwrnote) = True Then
                                        GoTo SkipforNull
                                        End If
                                        If IsNull(SEPTEXT2(!hdwrnote, 2, 2, 48)) = True Or !hdwrnote = "" Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            'varAttributes2(iatHW).TextString = Left(!HDWRDESC, 22)
                                            varAttributes2(iatHW).textString = SEPTEXT2(!hdwrnote, 2, 2, 48)
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XDESC1" Then
                                    Debug.Print "*******>>>>> : " & varAttributes2(iatHW).TagString
                                    Debug.Print "*******>>>>> : " & !HDWRDESC
                                    
                                        If IsNull(!HDWRDESC) = True Then
                                        GoTo SkipforNull
                                        End If
                                        If IsNull(SEPTEXT2(!HDWRDESC, 1, 2, 48)) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            'varAttributes2(iatHW).TextString = Left(!HDWRDESC, 22)
                                            varAttributes2(iatHW).textString = SEPTEXT2(!HDWRDESC, 1, 2, 48)
                                            'varAttributes2(iatHW).TextString = Left(!HDWRDESC, 32)
                                        End If
                                    End If
                                    If varAttributes2(iatHW).TagString = "XDESC2" Then
                                        If IsNull(!HDWRDESC) = True Then
                                        GoTo SkipforNull
                                        End If
                                        If IsNull(SEPTEXT2(!HDWRDESC, 2, 2, 48)) = True Then
                                            varAttributes2(iatHW).textString = ""
                                        Else
                                            'varAttributes2(iatHW).TextString = Left(!HDWRDESC, 22)
                                            varAttributes2(iatHW).textString = SEPTEXT2(!HDWRDESC, 2, 2, 48)
                                            'varAttributes2(iatHW).TextString = Left(!HDWRDESC, 32)
                                        End If
                                    End If
                                
                                
                                '**************************
SkipforNull:
                                Next iatHW
                            
                            TblPT(1) = TblPT(1) - nJumpRow
                            .MoveNext
                            Next k
                            End With 'hdwr rows
                   RsHDwrSet.Close
                'attributeObj.Update
        
                '************************************************
                'txtRow = !drgno & Chr(9) & !oldsize
                'Set textObj = ThisDrawing.ModelSpace.AddText(txtRow, tblPT, txtHT)
                '******************************************************
                'textObj.StyleName = "cn"
                'textObj.Update
                Debug.Print "iathw :" & iatHW
        TblPT(1) = TblPT(1)
        .MoveNext
        
    Next i
End With
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                             (TblPT, nFootDwg, 1, 1, 1, 0)
''ThisDrawing.SelectionSets.Item("hdwrsched").Delete
RstHDProjSet.Close

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


ExitMKHDWRimg:
frmDoor.Hide

End Sub


Public Sub Start()
  'InsertADrawing "Z:\Blocks\C_Paper.dwg", "Model"
End Sub
'############################################################################
'############################################################################
'#
'#
'#      writeWinImg     2014-09-08
'#
'#
'#
'############################################################################
'############################################################################
Public Sub writeWinImg()
Dim DB As Database
Dim RstProjWin, RstProjWinItem As Recordset
Dim strProjWin, strProjWinItem, txtRow As String
Dim rnProjWin As Integer   'ProjSet no.
Dim rnProjWinItem As Integer    'EachSet No.
'Dim rcArray As Variant
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
Dim height As Double
Dim mode As Long
Dim prompt As String

'**********IMG********************
Dim nHeadDwg, nGroupDwg, nRowDwg, nFootDwg, nBubdwg As Variant
Dim nJumpHead, nJumpGroup, nJumpRow, rn As Integer
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

Dim xProjWinID, i, iatProjWin, k, iatProjWinItem As Integer
frmDoor.ListBox1.Clear
CreateCNstyle
'############################################################################
'#         Set DB
'############################################################################

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\Door02_01.mdb")

'############################################################################
'#         Set Points
'############################################################################

TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify
'############################################################################
'#         Set DWG List
'############################################################################
If frmDoor.writeWindImg = True Then
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XWINSCHED_HD02.dwg"
    'nGroupDwg = "Z:\jvba\XHDWRGPROW03.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\XWINSCHED_ROW02.dwg"
    'nFootDwg = "H:\drawfile\jvba\XHDWRFOOT03.dwg"
    nBubdwg = "\\jwja-svr-10\drawfile\jvba\XWINSCHED_BUB01.dwg"
    nJumpHead = 72 'was 24
    'nJumpGroup = 24 'was 66
    nJumpRow = 72
    'InsertJBlock "H:\drawfile\jvba\XHDWRHEAD03.dwg", tblPT, "Model"
ElseIf frmDoor.writeWindShort = True Then
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XWINSCHED_HD01.dwg"
    'nGroupDwg = "H:\drawfile\jvba\XHDWRGPROW02.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\XWINSCHED_ROW01.dwg"
    'nFootDwg = "H:\drawfile\jvba\XHDWRFOOT02.dwg"
    nBubdwg = "\\jwja-svr-10\drawfile\jvba\XWINSCHED_BUB01.dwg"
    nJumpHead = 48
    'nJumpGroup = 36
    nJumpRow = 16
    'InsertJBlock "H:\drawfile\jvba\XHDWRHEAD02.dwg", tblPT, "Model"
ElseIf frmDoor.writeWindLong = True Then
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XWINDSCHED_HD03.dwg"
    'nGroupDwg = "H:\drawfile\jvba\XHDWRGPROW02.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\XWINDSCHED_ROW03.dwg" 'XWINDSCHED_ROW03
    'nFootDwg = "H:\drawfile\jvba\XHDWRFOOT02.dwg"
    nBubdwg = "\\jwja-svr-10\drawfile\jvba\XWINSCHED_BUB01.dwg"
    nJumpHead = 60 '5x12
    'nJumpGroup = 36
    nJumpRow = 24 '2x12
    'InsertJBlock "H:\drawfile\jvba\XHDWRHEAD02.dwg", tblPT, "Model"

End If
'############################################################################
'#         rstProjWin
'############################################################################
    strProjWin = "SELECT projwin.projwinid, projwin.proj_no, projwin.connprojwin, projwin.PROJWINDNAME, projwin.note, projwin.input " _
    & "FROM projwin " _
    & "WHERE (((projwin.proj_no) = '" & frmDoor.seleproj & "')) " _
    & "ORDER BY projwin.projwinid;"
RecordSets:
    Set RstProjWin = DB.OpenRecordset(strProjWin, dbOpenDynaset)

If RstProjWin.RecordCount = 0 Then
    GoTo ExitWindImg
End If
'RstHDProjSet.MoveLast
'RstHDProjSet.MoveFirst
rnProjWin = RstProjWin.RecordCount
Debug.Print rnProjWin
'############################################################################
'#         Write Head
'############################################################################

With RstProjWin   'WITH 1
     xProjWinID = !projwinid
     Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                 (TblPT, nHeadDwg, 1, 1, 1, 0)
    'Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
           (insertionPnt, "xdrschedrow01", 1#, 1#, 1#, 0)
    'set blcokrefobj = thisdrawing.modelspace.Name.ObjectName "xdrschedrow01"
    '******************************
    'Get the block's attributes
    Dim varAttributes0 As Variant
    varAttributes0 = blockRefObj.GetAttributes
    
    For iatProjWin = LBound(varAttributes0) To UBound(varAttributes0)
    '############################################################################
    '#         XPROJ_NO
    '############################################################################
    Debug.Print varAttributes0(iatProjWin).TagString
    If varAttributes0(iatProjWin).TagString = "XPROJ_NO" Then
        If IsNull(!PROJ_NO) = True Then
            varAttributes0(iatProjWin).textString = ""
        Else
            varAttributes0(iatProjWin).textString = !PROJ_NO
            Debug.Print varAttributes0(iatProjWin).textString
        End If
    End If
    '############################################################################
    '#         XDATE
    '############################################################################
    Debug.Print varAttributes0(iatProjWin).TagString
    If varAttributes0(iatProjWin).TagString = "XDATE" Then
        'If IsNull(!Proj_no) = True Then
            'varAttributes0(iatProjWin).TextString = ""
        'Else
            varAttributes0(iatProjWin).textString = Date
            Debug.Print varAttributes0(iatProjWin).textString
        'End If
    End If
    '############################################################################
    '#         XPROJ_WINDSCHEDNAME
    '############################################################################
    Debug.Print varAttributes0(iatProjWin).TagString
    If varAttributes0(iatProjWin).TagString = "XPROJ_WINDSCHEDNAME" Then
        If IsNull(!PROJWINDNAME) = True Then
            varAttributes0(iatProjWin).textString = ""
        Else
            varAttributes0(iatProjWin).textString = !PROJWINDNAME
            Debug.Print varAttributes0(iatProjWin).textString
        End If
    End If
    
    Next iatProjWin
    TblPT(1) = TblPT(1) - nJumpHead

    '############################################################################
    '#         Write Rows
    '############################################################################
    strProjWinItem = "SELECT projwin.projwinid, connprojwin.connprojwinid, connprojwin.projwinno, connprojwin.projwinnoTE, " _
    & "WINDOW.WINSIZE, WINDOW.CATNO, WINDOW.WINMAT, WINDOW.WINTYPE, WINDOW.WINMODEL, WINDOW.WINAREA, window.winmfgr, " _
    & "WINDOW.WINFIN, WINDOW.WINGL, WINDOW.WINNOTE, WINDOW.WINOPENAREA, WINDOW.WINIMGNAME, WINDOW.WINGL, WINDOW.UFACTOR, WINDOW.SHGC, WINDOW.VT, WINDOW.AIRLEAK, " _
    & "Contact.COMPANY " _
    & "FROM ((projwin INNER JOIN connprojwin ON projwin.projwinid = connprojwin.projwinid) " _
    & "INNER JOIN WINDOW ON connprojwin.winid = WINDOW.WINID) " _
    & "INNER JOIN Contact ON WINDOW.WINMFGR = Contact.CID " _
    & "WHERE (((projwin.projwinid)=" & xProjWinID & ")) " _
    & "ORDER BY connprojwin.PROJWINNO;"
    
    '& "FROM (projwin INNER JOIN connprojwin ON projwin.projwinid = connprojwin.projwinid) " _
    '& "INNER JOIN WINDOW ON connprojwin.winid = WINDOW.WINID " _

    
    Set RstProjWinItem = DB.OpenRecordset(strProjWinItem)
    rnProjWinItem = RstProjWinItem.RecordCount
    Debug.Print rnProjWinItem
    RstProjWinItem.MoveLast
    RstProjWinItem.MoveFirst
    With RstProjWinItem 'With 2
        
        'Do Until RstProjWinItem.EOF = True
        For k = 1 To rnProjWinItem
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                        (TblPT, nRowDwg, 1, 1, 1, 0)
        Dim varAttributes As Variant
        varAttributes = blockRefObj.GetAttributes
                        
                        
            Debug.Print "Win No.: " & !projwinno
            '############################################################################
            '#         Image
            '############################################################################
            If !WINIMGNAME <> "" Then
                'If frmDoor.showWinImg = True Then
                '############################################################################
                '#         EDITED 18/10/28
                '############################################################################
                If frmDoor.writeWindImg = True Then
                    Dim imgPT(0 To 2) As Double
                    Dim scalefactor, xSF As Double
                    Dim rotAngleInDegree As Double, rotAngle As Double
                    'Dim imageName As String
                    Dim imageName As String
                
                    Dim raster As AcadRasterImage
                    Dim imgheight As Variant
                    Dim imgwidth As Variant
                
                    'imageName = "C:\AutoCAD\sample\downtown.jpg"
                    imageName = !WINIMGNAME
                    Debug.Print imageName
                    'insertionPoint(0) = 2#: insertionPoint(1) = 2#: insertionPoint(2) = 0#
                    'imgPT(0) = tblPT(0) + 36: imgPT(1) = tblPT(1) - 54: imgPT(2) = tblPT(2)
                    'imgPT(0) = tblPT(0) + 6: imgPT(1) = tblPT(1) - 54: imgPT(2) = tblPT(2)
                    imgPT(0) = TblPT(0) + 42: imgPT(1) = TblPT(1) - 60: imgPT(2) = TblPT(2)
                    'scalefactor = 48#
                    scalefactor = 60#
                    'rotAngleInDegree = 0#
                    'rotAngle = rotAngleInDegree * 3.141592 / 180#
                    rotAngle = 0#
                    
                    
                    
                    ' Creates a raster image in model space
                    'If !fximgname <> "" Then
                    Set raster = ThisDrawing.ModelSpace.AddRaster(imageName, imgPT, scalefactor, rotAngle)
                    With raster
                    imgheight = raster.ImageHeight
                    imgwidth = raster.ImageWidth
                        If imgheight > 60 Then
                            xSF = (60 / imgheight) * 60
                            Debug.Print "xSF: " & xSF * 60
                             raster.scalefactor = xSF
                        End If
                    End With
                End If
            End If
            '***************************************************************

            
            
            For iatProjWinItem = LBound(varAttributes) To UBound(varAttributes)
            Debug.Print varAttributes(iatProjWinItem).TagString
                '############################################################################
                '#         XWINNO
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINNO" Then
                    If IsNull(!projwinno) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !projwinno
                    '############################################################################
                    '#         WINBUBBLE
                    '############################################################################
                        
                        Dim varAttributes2 As Variant
                        Dim iatBub As Integer
                        TblPT(0) = TblPT(0) - 48
                        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                        (TblPT, nBubdwg, 1, 1, 1, 0)
                        varAttributes2 = blockRefObj.GetAttributes
                        For iatBub = LBound(varAttributes2) To UBound(varAttributes2)
                        Debug.Print iatBub
                            If varAttributes2(iatBub).TagString = "WINNO" Then
                                If IsNull(!projwinno) = True Then
                                    varAttributes2(iatBub).textString = ""
                                Else
                                    varAttributes2(iatBub).textString = !projwinno
                                End If
                            End If
                        Next iatBub
                        TblPT(0) = TblPT(0) + 48
                    '############################################################################
                    End If
                End If
                '############################################################################
                '#         XWINSIZE
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINSIZE" Then
                    If IsNull(!WINSIZE) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !WINSIZE
                    End If
                End If
                '############################################################################
                '#         XWINTYPE1
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINTYPE1" Then
                    If IsNull(!WINTYPE) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = SEPTEXT2(!WINTYPE, 1, 2, 12)
                    End If
                End If
                '############################################################################
                '#         XWINTYPE2
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINTYPE2" Then
                    If IsNull(!WINTYPE) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = SEPTEXT2(!WINTYPE, 2, 2, 12)
                    End If
                End If
                
                '############################################################################
                '#         XWINCAT
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINCAT" Then
                    If IsNull(!CATNO) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !CATNO
                    End If
                End If
                '############################################################################
                '#         XWINMAT
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINMAT" Then
                    If IsNull(!WINMAT) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !WINMAT
                    End If
                End If
                '############################################################################
                '#         XWINMODEL
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINMODEL" Then
                    If IsNull(!WINMODEL) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !WINMODEL
                    End If
                End If
                '############################################################################
                '#         XWINFIN
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINFIN" Then
                    If IsNull(!WINFIN) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !WINFIN
                    End If
                End If
                '############################################################################
                '#         XWINGL
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINGL" Then
                    If IsNull(!WINGL) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !WINGL
                    End If
                End If
                '############################################################################
                '#         XWINNO
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINNO" Then
                    If IsNull(!projwinno) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !projwinno
                    End If
                End If
                '############################################################################
                '#         XWINMFGR
                '############################################################################
                
                
                If varAttributes(iatProjWinItem).TagString = "XWINMFGR" Then
                    If IsNull(!COMPANY) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        
                        varAttributes(iatProjWinItem).textString = !COMPANY
                        'varAttributes(iatProjWinItem).TextString = DLookUp("company", "contact", "[cid]= 1122")
                    End If
                End If
                '############################################################################
                '#         XWINNOTE
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINNOTE" Then
                    If IsNull(!WINNOTE) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !WINNOTE
                    End If
                End If
                '############################################################################
                '#         XWINREMARK1
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINREMARK1" Then
                    If IsNull(!projwinnote) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = SEPTEXT2(!projwinnote, 1, 2, 18)
                    End If
                End If
                '############################################################################
                '#         XWINREMARK2
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINREMARK2" Then
                    If IsNull(!projwinnote) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = SEPTEXT2(!projwinnote, 2, 2, 18)
                    End If
                End If
                
                
                
                '############################################################################
                '#         XWINREMARK
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWINREMARK" Then
                    If IsNull(!projwinnote) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !projwinnote
                    End If
                End If
                
                
                
                '############################################################################
                '#         XU
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XU" Then
                    If IsNull(!UFACTOR) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !UFACTOR
                    End If
                End If
                '############################################################################
                '#         SHGC
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "SHGC" Then
                    If IsNull(!SHGC) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !SHGC
                    End If
                End If
                '############################################################################
                '#         XWVT
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWVT" Then
                    If IsNull(!VT) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !VT
                    End If
                End If
                '############################################################################
                '#         XWAIR
                '############################################################################
                If varAttributes(iatProjWinItem).TagString = "XWAIR" Then
                    If IsNull(!AIRLEAK) = True Then
                        varAttributes(iatProjWinItem).textString = ""
                    Else
                        varAttributes(iatProjWinItem).textString = !AIRLEAK
                    End If
                End If
            
            
            
            
            
            
            
            Next iatProjWinItem
            TblPT(1) = TblPT(1) - nJumpRow
            .MoveNext
            Debug.Print "K: " & k
        Next k
        'tblPT(1) = tblPT(1) - nJumpRow
        '.MoveNext
    End With 'With 2
    


End With    'WITH 1
ExitWindImg:
RstProjWin.Close
RstProjWinItem.Close

Set DB = Nothing
frmDoor.Hide

End Sub
