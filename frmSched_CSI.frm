VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSched_CSI 
   Caption         =   "frmSched_CSI"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   OleObjectBlob   =   "frmSched_CSI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSched_CSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Public Sub UserForm_Initialize()
Dim WS As Workspace
Dim DB As Database
Dim rstProj As Recordset
Dim strProj As String

' Define the workgroup information file (system database) you will use.
'DBEngine.SystemDB = "C:\Documents and Settings\Jinwoo\Application Data\Microsoft\Access\System6.mdw"
'DBEngine.SystemDB = "C:\Documents and Settings\jinwoo\Application Data\Microsoft\Access\secured.mdw"
'DBEngine.SystemDB = "C:\Documents and Settings\jwjang\Application Data\Microsoft\Access\system.mdw"
'DBEngine.SystemDB = "X:\JW0610.mdw"
DBEngine.SystemDB = "\\jwja-svr-10\jdb\JW0610.mdw"

' OPTIONAL: Create a workspace with the correct login information
'Set WS = CreateWorkspace("NewWS", "MyID", "MyPwd", dbUseJet)
' OPTIONAL: If you set the workspace object, you can use the new
'           workspace with the OpenDatabase command. For example:
'Set Db = WS.OpenDatabase("C:\My Documents\MyDB.mdb")
'
'Set DB = OpenDatabase("h:\db\db_07\JPM02_01.mdb")
'Set DB = OpenDatabase("h:\db\db_07\H_H_JPM02_01.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\DATAS\DT11_JPM_01.mdb")

'Set Rs = DB.OpenRecordset("project")
'SELECT Project.PROJ_NO, Project.PROJ_NAME, Project.STATUS, Project.INPUT
'FROM Project
'WHERE (((Project.Status) = "act"))
'ORDER BY Project.INPUT DESC;

strProj = "SELECT Project.PROJ_NO, Project.PROJ_NAME, Project.STATUS, Project.INPUT " _
& "FROM project " _
& "where (((Project.Status) = 'act'))" _
& "ORDER BY Project.INPUT DESC;"
Set rstProj = DB.OpenRecordset(strProj)
rstProj.MoveLast
rstProj.MoveFirst
Debug.Print rstProj.RecordCount
Do Until rstProj.EOF = True
    seleproj.AddItem rstProj!PROJ_NO
    seleproj.List(seleproj.ListCount - 1, 1) = rstProj!PROJ_NAME
    rstProj.MoveNext
Loop
rstProj.Close
Set rstProj = Nothing
DB.Close
Set DB = Nothing

' OPTIONAL: If you set the WS object, use the following to clean up
'WS.Close
'Set WS = Nothing

MsgBox "Projects Loaded!"

End Sub
Private Sub CheckBox1_Click()

End Sub

Private Sub CommandButton1_Click()
Unload frmSched_CSI
'frmSched_CSI.HIDE

End Sub
Private Sub listSubmission_Click()
'FillSubmission
Dim DB As Database
Dim RstDwg As Recordset
Dim strDwg As String

ListBox4.Clear 'dwgitems

'========================================================================================
' Retired Qry 9/21/18
'========================================================================================
strDwg_old = "SELECT TRADE.TRADENO, DWGmain.ID, DWGmain.TRADE, DWGmain.DWGGROUP, " _
& "Trim([trade] & [dwggroup] & ' ' & [dwgno]) AS dwgnumber, DWGmain.SECTOR, " _
& "DWGmain.DWGNO, DWGmain.DWGNAME, DWGmain.DATE, DWGmain.REV, DWGmain.COMPLETE, " _
& "DWGmain.CREATEDBY, DWGmain.GENnote, DWGmain.taskid, DWGmain.NOTEID, DWGGROUP.GROUPNAME, DWGmain.[NO], DWGmain.proj_no, TRADE.TRADENAME " _
& "FROM TRADE INNER JOIN (DWGmain LEFT JOIN DWGGROUP ON DWGmain.DWGGROUP = DWGGROUP.GROUPNO) ON TRADE.TRADEKEY = DWGmain.TRADE " _
& "WHERE (((DWGmain.PROJ_NO) ='" & seleproj & "')) " _
& "ORDER BY TRADE.TRADENO, DWGmain.TRADE, DWGmain.DWGGROUP, Trim([trade] & [dwggroup] & ' ' & [dwgno]);"

'========================================================================================
' New Qry 9/21/18
'========================================================================================
strDwg = "SELECT TRADE.TRADENO, DWGmain.ID, DWGmain.TRADE, DWGmain.DWGGROUP, " _
& "Trim([trade] & [dwggroup] & ' ' & [dwgno]) AS dwgnumber, DWGmain.SECTOR, " _
& "DWGmain.DWGNO, DWGmain.DWGNAME, DWGmain.DATE, DWGmain.REV, DWGmain.COMPLETE, " _
& "DWGmain.CREATEDBY, DWGmain.GENnote, DWGmain.taskid, DWGmain.NOTEID, DWGGROUP.GROUPNAME, " _
& "DWGmain.[NO], DWGmain.proj_no, TRADE.TRADENAME, SUBMISSION.SUBID " _
& "FROM (TRADE INNER JOIN (DWGmain LEFT JOIN DWGGROUP ON DWGmain.DWGGROUP = DWGGROUP.GROUPNO) " _
& "ON TRADE.TRADEKEY = DWGmain.TRADE) INNER JOIN (SUBMISSION INNER JOIN SubItem ON SUBMISSION.SUBID = SubItem.subid) ON DWGmain.ID = SubItem.subdwg " _
& "WHERE (((SUBMISSION.SUBID) = " & listSubmission & ")) " _
& "ORDER BY TRADE.TRADENO, DWGmain.TRADE, DWGmain.DWGGROUP, Trim([trade] & [dwggroup] & ' ' & [dwgno]);"

strDwg_test = "SELECT TRADE.TRADENO, DWGmain.ID, DWGmain.TRADE, DWGmain.DWGGROUP, " _
& "Trim([trade] & [dwggroup] & ' ' & [dwgno]) AS dwgnumber, DWGmain.SECTOR, " _
& "DWGmain.DWGNO, DWGmain.DWGNAME, DWGmain.DATE, DWGmain.REV, DWGmain.COMPLETE, " _
& "DWGmain.CREATEDBY, DWGmain.GENnote, DWGmain.taskid, DWGmain.NOTEID, DWGGROUP.GROUPNAME, " _
& "DWGmain.[NO], DWGmain.proj_no, TRADE.TRADENAME "

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\X-dwg.mdb")

Set RstDwg = DB.OpenRecordset(strDwg)

Do Until RstDwg.EOF = True
    ListBox4.AddItem RstDwg!DWGNUMBER
    If IsNull(RstDwg!DWGNAME) = False Then
    ListBox4.List(ListBox4.ListCount - 1, 1) = RstDwg!DWGNAME
    End If
    RstDwg.MoveNext
Loop
RstDwg.Close
Set RstDwg = Nothing


DB.Close
Set DB = Nothing

End Sub

Private Sub seleproj_Change()
'===================================================================================
' DWG
'===================================================================================

Dim DB As Database
Dim RstDwg As Recordset
Dim strDwg As String
'DBEngine.SystemDB = "C:\Documents and Settings\Jinwoo\Application Data\Microsoft\Access\secured.mdw"

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\X-dwg.mdb")
'------------------------------------------------------------------------

listSubmission.Clear
strSub = "SELECT SUBMISSION.SUBID, SUBMISSION.SUBDATE, SUBMISSION.SUBNO, SUBMISSION.SUBNAME, SUBMISSION.PROJ_NO " _
& "FROM SUBMISSION " _
& "WHERE (((SUBMISSION.PROJ_NO)='" & seleproj & "'))"
Set RstSub = DB.OpenRecordset(strSub)
Do Until RstSub.EOF = True
    listSubmission.AddItem RstSub!SUBID
    If IsNull(RstSub!subname) = False Then
    'RstSub.AddItem RstSub!subid
    listSubmission.List(listSubmission.ListCount - 1, 1) = RstSub!subname
    End If
    RstSub.MoveNext
Loop
RstSub.Close
Set RstSub = Nothing
'------------------------------------------------------------------------
ListBox1.Clear 'dwgitems

'========================================================================================
' Retired Qry 9/21/18
'========================================================================================
strDwg = "SELECT TRADE.TRADENO, DWGmain.ID, DWGmain.TRADE, DWGmain.DWGGROUP, " _
& "Trim([trade] & [dwggroup] & ' ' & [dwgno]) AS dwgnumber, DWGmain.SECTOR, " _
& "DWGmain.DWGNO, DWGmain.DWGNAME, DWGmain.DATE, DWGmain.REV, DWGmain.COMPLETE, " _
& "DWGmain.CREATEDBY, DWGmain.GENnote, DWGmain.taskid, DWGmain.NOTEID, DWGGROUP.GROUPNAME, DWGmain.[NO], DWGmain.proj_no, TRADE.TRADENAME " _
& "FROM TRADE INNER JOIN (DWGmain LEFT JOIN DWGGROUP ON DWGmain.DWGGROUP = DWGGROUP.GROUPNO) ON TRADE.TRADEKEY = DWGmain.TRADE " _
& "WHERE (((DWGmain.PROJ_NO) ='" & seleproj & "')) " _
& "ORDER BY TRADE.TRADENO, DWGmain.TRADE, DWGmain.DWGGROUP, Trim([trade] & [dwggroup] & ' ' & [dwgno]);"

'========================================================================================
' New Qry 9/21/18
'========================================================================================
strDwg_New = "SELECT TRADE.TRADENO, DWGmain.ID, DWGmain.TRADE, DWGmain.DWGGROUP, " _
& "Trim([trade] & [dwggroup] & ' ' & [dwgno]) AS dwgnumber, DWGmain.SECTOR, " _
& "DWGmain.DWGNO, DWGmain.DWGNAME, DWGmain.DATE, DWGmain.REV, DWGmain.COMPLETE, " _
& "DWGmain.CREATEDBY, DWGmain.GENnote, DWGmain.taskid, DWGmain.NOTEID, DWGGROUP.GROUPNAME, " _
& "DWGmain.[NO], DWGmain.proj_no, TRADE.TRADENAME, SUBMISSION.SUBID " _
& "FROM (TRADE INNER JOIN (DWGmain LEFT JOIN DWGGROUP ON DWGmain.DWGGROUP = DWGGROUP.GROUPNO) " _
& "ON TRADE.TRADEKEY = DWGmain.TRADE) INNER JOIN (SUBMISSION INNER JOIN SubItem ON SUBMISSION.SUBID = SubItem.subid) ON DWGmain.ID = SubItem.subdwg " _
& "WHERE (((SUBMISSION.SUBID) = '" & listSubmission & "')) " _
& "ORDER BY TRADE.TRADENO, DWGmain.TRADE, DWGmain.DWGGROUP, Trim([trade] & [dwggroup] & ' ' & [dwgno]);"



Set RstDwg = DB.OpenRecordset(strDwg)

Do Until RstDwg.EOF = True
    ListBox1.AddItem RstDwg!DWGNUMBER
    If IsNull(RstDwg!DWGNAME) = False Then
    ListBox1.List(ListBox1.ListCount - 1, 1) = RstDwg!DWGNAME
    End If
    RstDwg.MoveNext
Loop
RstDwg.Close
Set RstDwg = Nothing


DB.Close
Set DB = Nothing
'===================================================================================
' FIXTURE_NEW WITH [PROFIXLIST]
'===================================================================================
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_FIX\X-FIXTURE_17_01.mdb")
ListBox5.Clear
strProFixList = "SELECT PROFIXLIST.PROJ_NO, PROFIXLIST.PROFIXLISTID, PROFIXLIST.PROFIXLISTNAME, PROFIXLIST.PROFIXLISTDATE " _
& "FROM PROFIXLIST " _
& "WHERE (((PROFIXLIST.PROJ_NO)='" & seleproj & "')) " _
& "ORDER BY PROFIXLISTID;"
Set RstProFixList = DB.OpenRecordset(strProFixList)

Do Until RstProFixList.EOF = True
    If IsNull(RstProFixList!PROFIXLISTID) = False Then

    ListBox5.AddItem RstProFixList!PROFIXLISTID
    ListBox5.List(ListBox5.ListCount - 1, 1) = RstProFixList!PROFIXLISTNAME
    End If
    RstProFixList.MoveNext
Loop
RstProFixList.Close
Set RstProFixList = Nothing
DB.Close
Set DB = Nothing
'---------------------------------------------
' 11/3/18
'---------------------------------------------


'===================================================================================
' FIXTURE_OLD
'===================================================================================
'Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_FIX\X-FIXTURE_17_01.mdb")

'ListBox2.Clear
'strFXType = "SELECT FIXITEM.FXTYPE " _
'& "FROM FIXITEM INNER JOIN PROFIX ON FIXITEM.FXID = PROFIX.FXID " _
'& "GROUP BY PROFIX.PROJ_NO, FIXITEM.FXTYPE " _
'& "HAVING (((PROFIX.PROJ_NO) = '" & seleproj & "')) " _
'& "ORDER BY FIXITEM.FXTYPE;"

'Set RstFXType = DB.OpenRecordset(strFXType)

'Do Until RstFXType.EOF = True
'    If IsNull(RstFXType!FXTYPE) = False Then

'    ListBox2.AddItem RstFXType!FXTYPE
'    'ListBox1.List(ListBox1.ListCount - 1, 1) = RstFXType!DWGNAME
'    End If
'    RstFXType.MoveNext
'Loop
'RstFXType.Close
'Set RstFXType = Nothing
'DB.Close
'Set DB = Nothing
'===================================================================================
' LIGHT
'===================================================================================

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\X-light.mdb")

ListBox3.Clear
strLight = "SELECT PROFIX.PROJ_NO, PROFIX.PFXNO, FIXITEM.FXTYPE, FIXITEM.FXMODEL " _
& "FROM FIXITEM INNER JOIN PROFIX ON FIXITEM.FXID = PROFIX.FXID " _
& "WHERE (((PROFIX.PROJ_NO)='" & seleproj & "'))" _
& "ORDER BY PROFIX.PFXNO;"

'& "HAVING (((PROFIX.PROJ_NO) = '" & seleproj & "')) " _
'& "ORDER BY PROFIX.PFXNO;"

Set RstLight = DB.OpenRecordset(strLight)
Debug.Print RstLight.RecordCount
If RstLight.RecordCount = 0 Then
    Exit Sub
Else
x = RstLight.RecordCount
Do Until RstLight.EOF = True
    If IsNull(RstLight!FXtype) = False Then
    x = RstLight.RecordCount
        If RstLight!PFXNO = "" Then
            GoTo SKIPFORNEXTLIGHT
        Else
            ListBox3.AddItem RstLight!PFXNO
            ListBox3.List(ListBox3.ListCount - 1, 1) = RstLight!FXMODEL
            'ListBox3.List(ListBox3.ListCount - 1, 2) = RstLight!fx
        End If
    End If
SKIPFORNEXTLIGHT:
    RstLight.MoveNext
Loop
RstLight.Close
Set RstLight = Nothing
DB.Close
Set DB = Nothing
End If
End Sub
Private Sub ListBox5_Click()
'=====================================
' 11/3/18
'=====================================
ListBox2.Clear 'FXTYPE
Dim DB As Database
Dim RSTPROFIX As Recordset
'strFXType = "SELECT FIXITEM.FXTYPE " _
& "FROM FIXITEM INNER JOIN PROFIX ON FIXITEM.FXID = PROFIX.FXID " _
& "GROUP BY PROFIX.PROJ_NO, FIXITEM.FXTYPE " _
& "HAVING (((PROFIX.PROJ_NO) = '" & seleproj & "')) " _
& "ORDER BY FIXITEM.FXTYPE;"



strProFix = "SELECT PROFIX.PROFIXLISTID, FIXITEM.FXTYPE " _
& "FROM FIXITEM INNER JOIN PROFIX ON FIXITEM.FXID = PROFIX.FXID " _
& "GROUP BY PROFIX.PROFIXLISTID, FIXITEM.FXTYPE " _
& "HAVING (PROFIX.PROFIXLISTID) = " & ListBox5 _
& " ORDER BY FIXITEM.FXTYPE;"

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_FIX\X-FIXTURE_17_01.mdb")

Set RSTPROFIX = DB.OpenRecordset(strProFix)

Do Until RSTPROFIX.EOF = True
    If IsNull(RSTPROFIX!FXtype) = False Then

    ListBox2.AddItem RSTPROFIX!FXtype
    'ListBox2.List(ListBox2.ListCount - 1, 1) = RSTPROFIX!PFXNO
    End If
    RSTPROFIX.MoveNext
Loop
RSTPROFIX.Close
Set RSTPROFIX = Nothing
DB.Close
Set DB = Nothing

End Sub


Private Sub CommandButton2_Click() 'Old DWG List Button
Dim DB As Database
Dim RstDwg, RstTrade As Recordset
Dim strDwg, strTrade As String
Dim rcTrade, rcDwg As Integer
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

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\X-dwg.mdb")

strTrade = "SELECT TRADE.TRADENAME, TRADE.TRADENO, DWGmain.proj_no " _
& "FROM TRADE LEFT JOIN DWGmain ON TRADE.TRADEKEY = DWGmain.TRADE " _
& "GROUP BY TRADE.TRADENAME, TRADE.TRADENO, DWGmain.proj_no " _
& "HAVING (((DWGmain.proj_no)='" & seleproj & "'));"

Set RstTrade = DB.OpenRecordset(strTrade)

TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify

ThisDrawing.ModelSpace.InsertBlock TblPT, "\\jwja-svr-10\drawfile\jvba\XDWGHEAD01.dwg", 1, 1, 1, 0
InsertJBlock "\\jwja-svr-10\drawfile\jvba\XDWGTRADE01.dwg", tblPT2, "Model"
InsertJBlock "\\jwja-svr-10\drawfile\jvba\XDWGROW01.dwg", tblPT2, "Model"
'InsertJBlock "e:\drawfile\jvba\XHDWRFOOT02.dwg", tblPT2, "Model"
TblPT(1) = TblPT(1) - 36 'HEAD
Set RstTrade = DB.OpenRecordset(strTrade)
RstTrade.MoveLast
RstTrade.MoveFirst
rcTrade = RstTrade.RecordCount
Debug.Print rcTrade
With RstTrade
    For i = 1 To rcTrade
        ' xHDWRSETID = !HDWRSETID
         Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, "XDWGTRADE01", 1, 1, 1, 0)
            Dim varAttributes As Variant
            varAttributes = blockRefObj.GetAttributes
            For iatTrade = LBound(varAttributes) To UBound(varAttributes)
            
            Debug.Print varAttributes(iatTrade).TagString
                If varAttributes(iatTrade).TagString = "XTRADENAME" Then
                    If IsNull(!TRADENAME) = True Then
                        varAttributes(iatTrade).textString = ""
                    Else
                        varAttributes(iatTrade).textString = !TRADENAME
                        Debug.Print varAttributes(iatTrade).textString
                    End If

                End If
                If varAttributes(iatTrade).TagString = "XTRADENO" Then
                    If IsNull(!TRADENO) = True Then
                        varAttributes(iatTrade).textString = ""
                    Else
                        varAttributes(iatTrade).textString = !TRADENO
                    End If

                End If
            Next iatTrade
            
                TblPT(1) = TblPT(1) - 18 'TRADE
                xTRADENO = !TRADENO
                Debug.Print xTRADENO
'_____________________________________________________________________________________
                strDwg = "SELECT TRADE.TRADENO, DWGmain.ID, DWGmain.TRADE, DWGmain.DWGGROUP, " _
& "Trim([trade] & [dwggroup] & ' ' & [dwgno]) AS dwgnumber, DWGmain.SECTOR, " _
& "DWGmain.DWGNO, DWGmain.DWGNAME, DWGmain.DATE, DWGmain.REV, DWGmain.COMPLETE, " _
& "DWGmain.CREATEDBY, DWGmain.GENnote, DWGmain.taskid, DWGmain.NOTEID, DWGGROUP.GROUPNAME, DWGmain.NO, DWGmain.proj_no, TRADE.TRADENAME " _
& "FROM TRADE INNER JOIN (DWGmain LEFT JOIN DWGGROUP ON DWGmain.DWGGROUP = DWGGROUP.GROUPNO) ON TRADE.TRADEKEY = DWGmain.TRADE " _
& "WHERE (((DWGmain.PROJ_NO) ='" & seleproj & "')) and trade.tradeno =" & xTRADENO _
& " ORDER BY TRADE.TRADENO, DWGmain.TRADE, DWGmain.DWGGROUP, Trim([trade] & [dwggroup] & ' ' & [dwgno]);"

                Set RstDwg = DB.OpenRecordset(strDwg)
                Debug.Print RstDwg.RecordCount
                'Debug.Print strHDSet

                
                If RstDwg.RecordCount = 0 Then
                    GoTo EXITDWG
                End If
                RstDwg.MoveLast
                RstDwg.MoveFirst
                Debug.Print RstDwg.RecordCount
                Debug.Print rcDwg
                rcDwg = RstDwg.RecordCount
                With RstDwg
                Debug.Print xTRADENO
                    For k = 1 To rcDwg
                         Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                     (TblPT, "XDWGROW01", 1, 1, 1, 0)
                        Dim varAttributes2 As Variant
                        varAttributes2 = blockRefObj.GetAttributes
                        For iatDWG = LBound(varAttributes2) To UBound(varAttributes2)
                        Debug.Print varAttributes2(iatDWG).TagString
                            If varAttributes2(iatDWG).TagString = "XDWGNO" Then
                                If IsNull(!no) = True Then
                                    varAttributes2(iatDWG).textString = ""
                                Else
                                    varAttributes2(iatDWG).textString = !no
                                End If
                            End If
                            If varAttributes2(iatDWG).TagString = "XDWGNAME" Then
                                If IsNull(!DWGNAME) = True Then
                                    varAttributes2(iatDWG).textString = ""
                                Else
                                    varAttributes2(iatDWG).textString = !DWGNAME
                                End If
                            End If
                            If varAttributes2(iatDWG).TagString = "XDATE" Then
                                If IsNull(!Date) = True Then
                                    varAttributes2(iatDWG).textString = ""
                                Else
                                    varAttributes2(iatDWG).textString = !Date
                                End If
                            End If
                            If varAttributes2(iatDWG).TagString = "XREV" Then
                                If IsNull(!REV) = True Then
                                    varAttributes2(iatDWG).textString = ""
                                Else
                                    varAttributes2(iatDWG).textString = !REV
                                End If
                            End If
                            If varAttributes2(iatDWG).TagString = "XNOTE1" Then
                                If IsNull(!GENnote) = True Then
                                    varAttributes2(iatDWG).textString = ""
                                Else
                                    varAttributes2(iatDWG).textString = !GENnote
                                End If
                            End If
                        Next iatDWG
                        TblPT(1) = TblPT(1) - 12 'ROW
                        .MoveNext
                    Next k
                    End With
                'tblPT(1) = tblPT(1) - 12
                .MoveNext
                Next i
                End With

EXITDWG:


RstTrade.Close

RstDwg.Close
Set rsttrad = Nothing
Set RstDwg = Nothing
DB.Close
Set DB = Nothing
frmSched_CSI.Hide
End Sub
Private Sub CommandButton5_Click() 'New DWG List Button
Dim DB As Database
Dim RstDwg, RstTrade As Recordset
Dim strDwg, strTrade As String
Dim rcTrade, rcDwg As Integer
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

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\X-dwg.mdb")

strTrade = "SELECT TRADE.TRADENAME, TRADE.TRADENO, DWGmain.proj_no " _
& "FROM TRADE LEFT JOIN DWGmain ON TRADE.TRADEKEY = DWGmain.TRADE " _
& "GROUP BY TRADE.TRADENAME, TRADE.TRADENO, DWGmain.proj_no " _
& "HAVING (((DWGmain.proj_no)='" & seleproj & "'));"

Set RstTrade = DB.OpenRecordset(strTrade)

TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify

ThisDrawing.ModelSpace.InsertBlock TblPT, "\\jwja-svr-10\drawfile\jvba\XDWGHEAD01.dwg", 1, 1, 1, 0
InsertJBlock "\\jwja-svr-10\drawfile\jvba\XDWGTRADE01.dwg", tblPT2, "Model"
InsertJBlock "\\jwja-svr-10\drawfile\jvba\XDWGROW01.dwg", tblPT2, "Model"
'InsertJBlock "e:\drawfile\jvba\XHDWRFOOT02.dwg", tblPT2, "Model"
TblPT(1) = TblPT(1) - 36 'HEAD
Set RstTrade = DB.OpenRecordset(strTrade)
RstTrade.MoveLast
RstTrade.MoveFirst
rcTrade = RstTrade.RecordCount
Debug.Print rcTrade
With RstTrade
    For i = 1 To rcTrade
        ' xHDWRSETID = !HDWRSETID
         Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, "XDWGTRADE01", 1, 1, 1, 0)
            Dim varAttributes As Variant
            varAttributes = blockRefObj.GetAttributes
            For iatTrade = LBound(varAttributes) To UBound(varAttributes)
            
            Debug.Print varAttributes(iatTrade).TagString
                If varAttributes(iatTrade).TagString = "XTRADENAME" Then
                    If IsNull(!TRADENAME) = True Then
                        varAttributes(iatTrade).textString = ""
                    Else
                        varAttributes(iatTrade).textString = !TRADENAME
                        Debug.Print varAttributes(iatTrade).textString
                    End If

                End If
                If varAttributes(iatTrade).TagString = "XTRADENO" Then
                    If IsNull(!TRADENO) = True Then
                        varAttributes(iatTrade).textString = ""
                    Else
                        varAttributes(iatTrade).textString = !TRADENO
                    End If

                End If
            Next iatTrade
            
                TblPT(1) = TblPT(1) - 18 'TRADE
                xTRADENO = !TRADENO
                Debug.Print xTRADENO
'_____________________________________________________________________________________
'==========================================================================================================================
' 9/21/18
'==========================================================================================================================

                strDwg_old = "SELECT TRADE.TRADENO, DWGmain.ID, DWGmain.TRADE, DWGmain.DWGGROUP, " _
& "Trim([trade] & [dwggroup] & ' ' & [dwgno]) AS dwgnumber, DWGmain.SECTOR, " _
& "DWGmain.DWGNO, DWGmain.DWGNAME, DWGmain.DATE, DWGmain.REV, DWGmain.COMPLETE, " _
& "DWGmain.CREATEDBY, DWGmain.GENnote, DWGmain.taskid, DWGmain.NOTEID, DWGGROUP.GROUPNAME, DWGmain.NO, DWGmain.proj_no, TRADE.TRADENAME " _
& "FROM TRADE INNER JOIN (DWGmain LEFT JOIN DWGGROUP ON DWGmain.DWGGROUP = DWGGROUP.GROUPNO) ON TRADE.TRADEKEY = DWGmain.TRADE " _
& "WHERE (((DWGmain.PROJ_NO) ='" & seleproj & "')) and trade.tradeno =" & xTRADENO _
& " ORDER BY TRADE.TRADENO, DWGmain.TRADE, DWGmain.DWGGROUP, Trim([trade] & [dwggroup] & ' ' & [dwgno]);"
'==========================================================================================================================
'strDwg = "SELECT TRADE.TRADENO, DWGmain.ID, DWGmain.TRADE, DWGmain.DWGGROUP, " _
& "Trim([trade] & [dwggroup] & ' ' & [dwgno]) AS dwgnumber, DWGmain.SECTOR, " _
& "DWGmain.DWGNO, DWGmain.DWGNAME, DWGmain.DATE, DWGmain.REV, DWGmain.COMPLETE, " _
& "DWGmain.CREATEDBY, DWGmain.GENnote, DWGmain.taskid, DWGmain.NOTEID, DWGGROUP.GROUPNAME, " _
& "DWGmain.[NO], DWGmain.proj_no, TRADE.TRADENAME, SUBMISSION.SUBID " _
& "FROM TRADE INNER JOIN (DWGmain LEFT JOIN DWGGROUP ON DWGmain.DWGGROUP = DWGGROUP.GROUPNO) " _
& "ON TRADE.TRADEKEY = DWGmain.TRADE) INNER JOIN (SUBMISSION INNER JOIN SubItem ON SUBMISSION.SUBID = SubItem.subid) ON DWGmain.ID = SubItem.subdwg " _
& "WHERE (((SUBMISSION.SUBID) = " & listSubmission & ")) And trade.tradeno =" & xTRADENO _
& "ORDER BY TRADE.TRADENO, DWGmain.TRADE, DWGmain.DWGGROUP, Trim([trade] & [dwggroup] & ' ' & [dwgno]);"
'==========================================================================================================================
strDwg_notworking = "SELECT TRADE.TRADENO, DWGmain.ID, DWGmain.TRADE, DWGmain.DWGGROUP, " _
& "Trim([trade] & [dwggroup] & ' ' & [dwgno]) AS dwgnumber, DWGmain.SECTOR, " _
& "DWGmain.DWGNO, DWGmain.DWGNAME, DWGmain.DATE, DWGmain.REV, DWGmain.COMPLETE, " _
& "DWGmain.CREATEDBY, DWGmain.GENnote, DWGmain.taskid, DWGmain.NOTEID, DWGGROUP.GROUPNAME, DWGmain.NO, DWGmain.proj_no, TRADE.TRADENAME " _
& "FROM TRADE INNER JOIN (DWGmain LEFT JOIN DWGGROUP ON DWGmain.DWGGROUP = DWGGROUP.GROUPNO) ON TRADE.TRADEKEY = DWGmain.TRADE " _
& "INNER JOIN (SUBMISSION INNER JOIN SubItem ON SUBMISSION.SUBID = SubItem.subid) ON DWGmain.ID = SubItem.subdwg " _
& "WHERE trade.tradeno =" & xTRADENO & " and (SUBMISSION.SUBID) = " & listSubmission & "))" _
& "ORDER BY TRADE.TRADENO, DWGmain.TRADE, DWGmain.DWGGROUP, Trim([trade] & [dwggroup] & ' ' & [dwgno]);"

strDwg = "SELECT TRADE.TRADENO, DWGmain.ID, DWGmain.TRADE, DWGmain.DWGGROUP, " _
& "Trim([trade] & [dwggroup] & ' ' & [dwgno]) AS dwgnumber, DWGmain.SECTOR, " _
& "DWGmain.DWGNO, DWGmain.DWGNAME, DWGmain.DATE, DWGmain.REV, DWGmain.COMPLETE, " _
& "DWGmain.CREATEDBY, DWGmain.GENnote, DWGmain.taskid, DWGmain.NOTEID, DWGGROUP.GROUPNAME, DWGmain.[NO], DWGmain.proj_no, TRADE.TRADENAME, SUBMISSION.SUBID " _
& "FROM (TRADE INNER JOIN (DWGmain LEFT JOIN DWGGROUP ON DWGmain.DWGGROUP = DWGGROUP.GROUPNO) ON TRADE.TRADEKEY = DWGmain.TRADE) " _
& "INNER JOIN (SUBMISSION INNER JOIN SubItem ON SUBMISSION.SUBID = SubItem.subid) ON DWGmain.ID = SubItem.subdwg " _
& "WHERE (((TRADE.TRADENO) =" & xTRADENO & ") And ((SUBMISSION.SUBID) =" & listSubmission & ")) " _
& "ORDER BY TRADE.TRADENO, DWGmain.TRADE, DWGmain.DWGGROUP, Trim([trade] & [dwggroup] & ' ' & [dwgno]);"

'==========================================================================================================================
' 9/21/18
'==========================================================================================================================
                Set RstDwg = DB.OpenRecordset(strDwg)
                Debug.Print RstDwg.RecordCount
                'Debug.Print strHDSet

                
                If RstDwg.RecordCount = 0 Then
                    GoTo EXITDWG
                End If
                RstDwg.MoveLast
                RstDwg.MoveFirst
                Debug.Print RstDwg.RecordCount
                Debug.Print rcDwg
                rcDwg = RstDwg.RecordCount
                With RstDwg
                Debug.Print xTRADENO
                    For k = 1 To rcDwg
                         Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                     (TblPT, "XDWGROW01", 1, 1, 1, 0)
                        Dim varAttributes2 As Variant
                        varAttributes2 = blockRefObj.GetAttributes
                        For iatDWG = LBound(varAttributes2) To UBound(varAttributes2)
                        Debug.Print varAttributes2(iatDWG).TagString
                            If varAttributes2(iatDWG).TagString = "XDWGNO" Then
                                If IsNull(!no) = True Then
                                    varAttributes2(iatDWG).textString = ""
                                Else
                                    varAttributes2(iatDWG).textString = !no
                                End If
                            End If
                            If varAttributes2(iatDWG).TagString = "XDWGNAME" Then
                                If IsNull(!DWGNAME) = True Then
                                    varAttributes2(iatDWG).textString = ""
                                Else
                                    varAttributes2(iatDWG).textString = !DWGNAME
                                End If
                            End If
                            If varAttributes2(iatDWG).TagString = "XDATE" Then
                                If IsNull(!Date) = True Then
                                    varAttributes2(iatDWG).textString = ""
                                Else
                                    varAttributes2(iatDWG).textString = !Date
                                End If
                            End If
                            If varAttributes2(iatDWG).TagString = "XREV" Then
                                If IsNull(!REV) = True Then
                                    varAttributes2(iatDWG).textString = ""
                                Else
                                    varAttributes2(iatDWG).textString = !REV
                                End If
                            End If
                            If varAttributes2(iatDWG).TagString = "XNOTE1" Then
                                If IsNull(!GENnote) = True Then
                                    varAttributes2(iatDWG).textString = ""
                                Else
                                    varAttributes2(iatDWG).textString = !GENnote
                                End If
                            End If
                        Next iatDWG
                        TblPT(1) = TblPT(1) - 12 'ROW
                        .MoveNext
                    Next k
                    End With
                'tblPT(1) = tblPT(1) - 12
                .MoveNext
                Next i
                End With

EXITDWG:


RstTrade.Close

RstDwg.Close
Set rsttrad = Nothing
Set RstDwg = Nothing
DB.Close
Set DB = Nothing
frmSched_CSI.Hide
End Sub


Sub LIGHT()
Dim DB As Database
Dim RstFix As Recordset
Dim strLight As String
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



Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\X-light.mdb")

strLight = "SELECT FIXITEM.FXTYPE, PROFIX.PFXID, FIXITEM.FXID, PROFIX.PFXNO, PROFIX.PROJ_NO, PROFIX.PFXRM, PROFIX.Pfxnote, " _
& "FIXITEM.FXMFGR, FIXITEM.FXMODEL, FIXITEM.FXCAT, FIXITEM.FXCOLOR, FIXITEM.FXIMGNAME, FIXITEM.FXLISTPRICE, PROFIX.FXDISCOUNT, " _
& "PROFIX.FXPRICE, PROFIX.FXTAX, FIXITEM.FXLEAD, FIXITEM.fxnote, Project.PROJ_NAME, Project.PROJ_ADDR, Project.PROJ_CITY, " _
& "Project.PROJ_STATE, Project.PROJ_ZIP, PROFIX.FXQTY, PROFIX.FXPRNOTE, Contact.COMPANY, LTSUP.LTBULB, LTSUP.LTBULBQTY, " _
& "LTSUP.LTVOLT, BULB.BUBTYPE, BULB.BUBMODEL, BULB.BUBCAT, BULB.BUBTEMP, BULB.BUBMFGR, BULB.BUBWATT, FIXITEM.FXMFGR, " _
& "Contact.COMPANY, Contact.STREET, Contact.CITY, Contact.STATE, Contact.ZIP, Contact.TEL_BUS, Trim([CONTACT].[CITY] & ', ' & [CONTACT].[STATE] & ' ' & [CONTACT].[ZIP]) AS SCSZ, " _
& "Contact_1.COMPANY, Contact_1.STREET, Contact_1.CITY, Contact_1.STATE, Contact_1.ZIP, Contact_1.TEL_BUS, Trim([CONTACT_1].[CITY] & ', ' & [CONTACT_1].[STATE] & ' ' & [CONTACT_1].[ZIP]) AS MCSZ " _
& "FROM ((((Contact RIGHT JOIN FIXITEM ON Contact.CID = FIXITEM.FXSUPPLIER) LEFT JOIN LTSUP ON FIXITEM.FXID = LTSUP.FXID) " _
& "LEFT JOIN BULB ON LTSUP.LTBULB = BULB.BUBID) LEFT JOIN Contact AS Contact_1 ON FIXITEM.FXMFGR = Contact_1.CID) " _
& "INNER JOIN (Project RIGHT JOIN PROFIX ON Project.PROJ_NO = PROFIX.PROJ_NO) ON FIXITEM.FXID = PROFIX.FXID " _
& "WHERE (((PROFIX.PROJ_NO) ='" & seleproj & "')) " _
& "ORDER BY FIXITEM.FXTYPE, PROFIX.PFXNO;"

'& "WHERE (((FIXITEM.fxtype) = [FORMS]![FRMFIX]![fxtype]) And ((PROFIX.PROJ_NO) ='" & seleproj & "')) " _





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
& "ORDER BY PROFIX.PFXID, FIXITEM.FXID, PROFIX.PFXNO;"

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


TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -500: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify
If ShowImage = True Then
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XLIGHTHEAD02.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\XLIGHTROW02.dwg"
    nJumpHead = 60
    nJumpRow = 60
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\XLIGHTHEAD02.dwg", tblPT2, "Model"
Else
    nHeadDwg = "\\jwja-svr-10\drawfile\jvba\XLIGHTHEAD01.dwg"
    nRowDwg = "\\jwja-svr-10\drawfile\jvba\XLIGHTROW01.dwg"
    nJumpHead = 48
    nJumpRow = 24
    InsertJBlock "\\jwja-svr-10\drawfile\jvba\XLIGHTHEAD01.dwg", tblPT2, "Model"

End If
'ThisDrawing.ModelSpace.InsertBlock tblPT, "H:\drawfile\jvba\XFIXHEAD01.dwg", 1, 1, 1, 0
'InsertJBlock "H:\drawfile\jvba\XDWGTRADE01.dwg", tblPT2, "Model"
'InsertJBlock nHeadDwg, tblPT2, "Model"
'InsertJBlock "e:\drawfile\jvba\XHDWRFOOT02.dwg", tblPT2, "Model"
'tblPT(1) = tblPT(1) - 48

Set RstFix = DB.OpenRecordset(strLight)
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
            varAttributes2(iatType).textString = !FXtype & " FIXTURE SCHEDULE"
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
If ShowImage = True Then
    Dim imgPT(0 To 2) As Double
    Dim scalefactor, xSF As Double
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
                If IsNull(SEPTEXT2(!FXMODEL, 1, 2, 16)) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = SEPTEXT2(!FXMODEL, 1, 2, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XMODEL2" Then 'XMODEL2
                If IsNull(SEPTEXT2(!FXMODEL, 2, 2, 16)) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = SEPTEXT2(!FXMODEL, 2, 2, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XNAME" Then 'XNAME
                If IsNull(!FXCAT) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !FXCAT
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "'XCAT" Then 'XCAT
                If IsNull(!FXCAT) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !FXCAT
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XFIN" Then 'XFIN
                If IsNull(!FXCOLOR) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !FXCOLOR
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XMCOMP" Then 'XMCOMP
                If IsNull(![CONTACT_1.COMPANY]) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = ![CONTACT_1.COMPANY]
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
                    varAttributes(iatFix).textString = ![CONTACT.COMPANY]
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
                    varAttributes(iatFix).textString = SEPTEXT2(!FXPRNOTE, 1, 3, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XNOTE2" Then 'XNOTE2
                If IsNull(!FXPRNOTE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = SEPTEXT2(!FXPRNOTE, 2, 3, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XNOTE3" Then 'XNOTE3
                If IsNull(!FXPRNOTE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = SEPTEXT2(!FXPRNOTE, 3, 3, 16)
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XBUBTYPE" Then 'XBUBTYPE
                If IsNull(!BUBTYPE) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !BUBTYPE
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If
            If varAttributes(iatFix).TagString = "XBUBCAT" Then 'XBUBCAT
                If IsNull(!BUBCAT) = True Then
                    varAttributes(iatFix).textString = ""
                Else
                    varAttributes(iatFix).textString = !BUBCAT
                    Debug.Print varAttributes(iatFix).textString
                End If
            End If

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

frmSched_CSI.Hide


End Sub
Private Sub CommandButton3_Click()
FIXTURE
End Sub

Private Sub CommandButton4_Click()
LIGHT
End Sub







Sub zbackup_image()
    '***************************************************************
        Dim ssImg As AcadSelectionSet
        Dim ftype(0) As Integer
        Dim fdata(0) As Variant
        Dim dxfcode As Variant
        Dim dxfdata As Variant
        Dim rasterObj As AcadRasterImage
        
        ftype(0) = 1: fdata(0) = "IMAGE" ' type 0 = object type, 2 = name
        dxfcode = ftype(0)
        dxfdata = fdata(0)

        
        
        For Each ssImg In ThisDrawing.SelectionSets
        If ssImg.Name = "IMAGESET" Then
        Exit For
        End If
        Next ssImg
        
        If ssImg Is Nothing Then
        Set ssImg = ThisDrawing.SelectionSets.Add("IMAGESET")
        Else
        ssImg.Clear
        End If
        
        ssImg.Select acSelectionSetAll, , , dxfcode, dxfdata
'Set obj = ThisDrawing.ModelSpace.GetObject("Cadimg")
        If ssImg.Count > 0 Then
        
            'Dim oEntity As AcadEntity
            'Dim objImage As AcadRasterImage
            Dim imgEntity As AcadEntity
            'Dim rasterObj As AcadRasterImage
            
            For Each imgEntity In ssImg
            Set rasterObj = imgEntity
            If (rasterObj.Transparency) Then
            'MsgBox "yeah"
            rasterObj.Transparency = False
            Else
            rasterObj.Transparency = True
            End If
            'rasterObj.ImageFile = !fximgname
            Next
            Else
            MsgBox "No images inserted."
            'ThisDrawing.ModelSpace.rasterObj.ImageFile = !FXIMGNAME
            'Debug.Print ssImg.Item(1).Name
        End If
        
        'ThisDrawing.Regen True
        ssImg.Delete
        Set ssImg = Nothing
        
        
        
    '***************************************************************

End Sub
