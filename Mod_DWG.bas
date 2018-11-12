Attribute VB_Name = "Mod_DWG"
Public Sub FillSubmission()
Dim DB As Database
Dim RstDwg As Recordset
Dim strDwg As String
Dim myForm As Object
'DBEngine.SystemDB = "C:\Documents and Settings\Jinwoo\Application Data\Microsoft\Access\secured.mdw"

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\X-dwg.mdb")
'------------------------------------------------------------------------
Set Screen.ActiveForm = myForm
myForm.listSubmission.Clear
'forms("frmSched_CSI").listSubmission.Clear
Screen.ActiveForm.[listSubmission].Clear
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
& "WHERE (((SUBMISSION.SUBID) = '" & listSubmission & "')) " _
& "ORDER BY TRADE.TRADENO, DWGmain.TRADE, DWGmain.DWGGROUP, Trim([trade] & [dwggroup] & ' ' & [dwgno]);"



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

