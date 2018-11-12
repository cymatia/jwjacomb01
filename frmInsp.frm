VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInsp 
   Caption         =   "frmInsp"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   OleObjectBlob   =   "frmInsp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

End Sub

Private Sub CommandButton10_Click()
writeEnergyTabularNW
End Sub

Private Sub CommandButton8_Click()

Dim x As Integer
'If IsNull(ListBox1) = True Then
'If ListBox1.Selected(1) = False Then
'MsgBox "Must select a Assembly Item from the list!"
'Exit Sub
'Else
'x = ListBox1
writeEnergyInsp
'frmWorkItem.HIDE
'End If


End Sub

Private Sub CommandButton9_Click()
'writeEnergyTabular
writeEnergyTabularXT
End Sub

Private Sub seleProjECC_Change()
EnergyList
End Sub



Public Sub UserForm_Initialize()
Dim DB As Database
Dim rstProj As Recordset
Dim strProj As String

' Define the workgroup information file (system database) you will use.
'DBEngine.SystemDB = "C:\Documents and Settings\Jinwoo\Application Data\Microsoft\Access\System6.mdw"
'DBEngine.SystemDB = "C:\Documents and Settings\jwjang\Application Data\Microsoft\Access\system.mdw"
'DBEngine.SystemDB = "H:\DB\JW0610.mdw"
DBEngine.SystemDB = "\\jwja-svr-10\jdb\JW0610.mdw"

' OPTIONAL: Create a workspace with the correct login information
'Set WS = CreateWorkspace("NewWS", "MyID", "MyPwd", dbUseJet)
' OPTIONAL: If you set the workspace object, you can use the new
'           workspace with the OpenDatabase command. For example:
'Set Db = WS.OpenDatabase("C:\My Documents\MyDB.mdb")
'
'Set DB = OpenDatabase("H:\db\db_misc\dob_inspections_01.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_DOB_INSPECTIONS\X-dob_inspections_02.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")

'Set Rs = DB.OpenRecordset("project")

'strPCSI = "SELECT PROJSPEC.PROJSPECID, PROJSPEC.PROJ_NO, PROJSPEC.NOTE, " _
& "PROJSPEC.PROJSPECDATE, PROJSPEC.NOTE, Project.PROJ_NAME " _
& "FROM PROJSPEC INNER JOIN Project ON PROJSPEC.PROJ_NO = Project.PROJ_NO " _
& "ORDER BY PROJSPEC.PROJ_NO"

'strProj = "SELECT projcode.projcodeid, PROJCODE.PROJ_NO, Project.PROJ_NAME " _
& "FROM Project INNER JOIN PROJCODE ON Project.PROJ_NO = PROJCODE.PROJ_NO " _
& "ORDER BY Project.INPUT DESC"

strProj = "SELECT PROJINSP.PROJINSPID, PROJINSP.PROJ_NO, PROJINSP.INPUT, PROJINSP.NOTE, " _
& "PROJINSP.REMARK, PROJINSP.CID, Project.PROJ_NAME, Project.PROJ_STNO, Project.PROJ_ADDR, " _
& "Project.PROJ_CITY, Project.PROJ_STATE, Project.PROJ_ZIP, Project.BIS " _
& "FROM Project INNER JOIN PROJINSP ON Project.PROJ_NO = PROJINSP.PROJ_NO " _
& "ORDER BY PROJINSP.INPUT desc;"

Set rstProj = DB.OpenRecordset(strProj)

Do Until rstProj.EOF = True
    seleproj.AddItem rstProj!PROJinspID
    seleproj.List(seleproj.ListCount - 1, 1) = rstProj!PROJ_NO
    seleproj.List(seleproj.ListCount - 1, 2) = rstProj!PROJ_NAME
    'seleproj.List(seleproj.ListCount - 1, 2) = rstProj!NOTE
    rstProj.MoveNext
Loop
rstProj.Close
Set rstProj = Nothing
DB.Close
Set DB = Nothing

' OPTIONAL: If you set the WS object, use the following to clean up
'WS.Close
'Set WS = Nothing

MsgBox "Project List for Inspection LOADED!"


End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub CommandButton7_Click()
Dim x As Integer
'If IsNull(ListBox1) = True Then
'If ListBox1.Selected(1) = False Then
'MsgBox "Must select a Assembly Item from the list!"
'Exit Sub
'Else
'x = ListBox1
writeInsp
'frmWorkItem.HIDE
'End If

End Sub



Private Sub seleproj_Change()
'frmCode.Repaint
CommandButton4_Click 'POPULATE CSI
Debug.Print seleproj
End Sub

Public Sub CommandButton4_Click() 'Test Populate CSI
Dim rstProjInsp As Recordset
Dim strProjInsp As String
Dim DB As Database
Dim rn As Integer
Dim rcArray As Variant
ListBox1.Clear
ListBox1.MultiSelect = fmMultiSelectExtended
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("H:\db\db_misc\JSPEC04_01.mdb")
'Set DB = OpenDatabase("H:\db\db_est\H_H_EST_2012_001.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_DOB_INSPECTIONS\X-dob_inspections_02.mdb")
'strPCSI2 = "SELECT PROJCSI.projcsiid, PROJCSI.PROJ_NO, MCSI.MCSI, MCSI.MCSINAME " _
& "FROM MCSI INNER JOIN PROJCSI ON MCSI.MCSIID = PROJCSI.MCSIID " _
& "WHERE (((PROJCSI.PROJ_NO) = '" & seleproj & "'))" _
& "ORDER BY MCSI.MCSI"

'strProjAss = "SELECT ASS.ASSID, ASS.ESTID, ASS.ASSDIV, ASS.ASSNO, ASS.ASSSUBDIV, ASS.ASSNAME, ASS.proj_no " _
& "FROM ASS " _
& "WHERE (((ASS.proj_no) = '" & seleproj & "'))" _
& "ORDER BY ASS.ASSDIV, ASS.ASSNO"

strProjInsp = "SELECT PROJINSPITEM.proj_no, PROJINSPITEM.projinspid, PROJINSPITEM.INPUT, " _
& "INSPECTION.INSPTYPE, INSPECTION.CODENO, INSPECTION.INSPNAME " _
& "FROM INSPECTION INNER JOIN PROJINSPITEM ON INSPECTION.INSPID = PROJINSPITEM.INSPID " _
& "WHERE (((PROJINSPITEM.projinspid) = " & seleproj & ")) " _
& "ORDER BY INSPECTION.INSPTYPE, INSPECTION.INSPID;"



Set rstProjInsp = DB.OpenRecordset(strProjInsp)
Debug.Print rn
Do Until rstProjInsp.EOF = True
    ListBox1.AddItem rstProjInsp!PROJinspID
    ListBox1.List(ListBox1.ListCount - 1, 1) = rstProjInsp!input
    ListBox1.List(ListBox1.ListCount - 1, 2) = rstProjInsp!INSPTYPE
    ListBox1.List(ListBox1.ListCount - 1, 3) = rstProjInsp!INSPNAME
    rstProjInsp.MoveNext
Loop
rstProjInsp.Close
Set rstProjInsp = Nothing
DB.Close
Set DB = Nothing
PopulateEnergyCombo
'EnergyList
End Sub
Public Sub PopulateEnergyCombo() 'Populate Energy ComboBox
Dim DB As Database
Dim rstProjECCList As Recordset
Dim strProjECCList As String

' Define the workgroup information file (system database) you will use.
'DBEngine.SystemDB = "H:\DB\JW0610.mdw"
DBEngine.SystemDB = "\\jwja-svr-10\jdb\JW0610.mdw"

' OPTIONAL: Create a workspace with the correct login information
'Set WS = CreateWorkspace("NewWS", "MyID", "MyPwd", dbUseJet)
' OPTIONAL: If you set the workspace object, you can use the new
'           workspace with the OpenDatabase command. For example:
'Set Db = WS.OpenDatabase("C:\My Documents\MyDB.mdb")
'
'Set DB = OpenDatabase("H:\db\db_misc\dob_inspections_01.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_DOB_INSPECTIONS\X-dob_inspections_02.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")


'strProj = "SELECT projcode.projcodeid, PROJCODE.PROJ_NO, Project.PROJ_NAME " _
& "FROM Project INNER JOIN PROJCODE ON Project.PROJ_NO = PROJCODE.PROJ_NO " _
& "ORDER BY Project.INPUT DESC"

'strProj = "SELECT PROJINSP.PROJINSPID, PROJINSP.PROJ_NO, PROJINSP.INPUT, PROJINSP.NOTE, " _
& "PROJINSP.REMARK, PROJINSP.CID, Project.PROJ_NAME, Project.PROJ_STNO, Project.PROJ_ADDR, " _
& "Project.PROJ_CITY, Project.PROJ_STATE, Project.PROJ_ZIP, Project.BIS " _
& "FROM Project INNER JOIN PROJINSP ON Project.PROJ_NO = PROJINSP.PROJ_NO;"
strProjECCList = "SELECT PROJECC.projeccid, PROJECC.note, PROJECC.proj_no " _
& "FROM PROJECC " _
& "WHERE (((PROJECC.proj_no)='" & seleproj.Column(1) & "'));"

Set rstProjECCList = DB.OpenRecordset(strProjECCList)

Do Until rstProjECCList.EOF = True
    seleProjECC.AddItem rstProjECCList!projeccid
    seleProjECC.List(seleProjECC.ListCount - 1, 1) = rstProjECCList!PROJ_NO
    seleProjECC.List(seleProjECC.ListCount - 1, 2) = rstProjECCList!NOTE
    'seleproj.List(seleproj.ListCount - 1, 2) = rstProj!NOTE
    rstProjECCList.MoveNext
Loop
rstProjECCList.Close
Set rstProjECCList = Nothing
DB.Close
Set DB = Nothing

' OPTIONAL: If you set the WS object, use the following to clean up
'WS.Close
'Set WS = Nothing

MsgBox "Project ECC List for Project LOADED!"


End Sub


Public Sub EnergyList() 'Populate Energy ListBox
Dim rstProjECC As Recordset
Dim strProjECC As String
Dim DB As Database
Dim rn As Integer
Dim rcArray As Variant
ListBox2.Clear
ListBox2.MultiSelect = fmMultiSelectExtended
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("H:\db\db_misc\JSPEC04_01.mdb")
'Set DB = OpenDatabase("H:\db\db_est\H_H_EST_2012_001.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_DOB_INSPECTIONS\X-dob_inspections_02.mdb")
'strPCSI2 = "SELECT PROJCSI.projcsiid, PROJCSI.PROJ_NO, MCSI.MCSI, MCSI.MCSINAME " _
& "FROM MCSI INNER JOIN PROJCSI ON MCSI.MCSIID = PROJCSI.MCSIID " _
& "WHERE (((PROJCSI.PROJ_NO) = '" & seleproj & "'))" _
& "ORDER BY MCSI.MCSI"

'strProjAss = "SELECT ASS.ASSID, ASS.ESTID, ASS.ASSDIV, ASS.ASSNO, ASS.ASSSUBDIV, ASS.ASSNAME, ASS.proj_no " _
& "FROM ASS " _
& "WHERE (((ASS.proj_no) = '" & seleproj & "'))" _
& "ORDER BY ASS.ASSDIV, ASS.ASSNO"

'strProjInsp = "SELECT PROJINSPITEM.proj_no, PROJINSPITEM.projinspid, PROJINSPITEM.INPUT, " _
& "INSPECTION.INSPTYPE, INSPECTION.CODENO, INSPECTION.INSPNAME " _
& "FROM INSPECTION INNER JOIN PROJINSPITEM ON INSPECTION.INSPID = PROJINSPITEM.INSPID " _
& "WHERE (((PROJINSPITEM.projinspid) = " & seleproj & ")) " _
& "ORDER BY INSPECTION.INSPTYPE, INSPECTION.INSPID;"

strProjECC = "SELECT PROJECC.projeccid, PROJECCITEM.projeccitemid, " _
& "PROJECCITEM.ecctabid, ECC.citation, ECC.provision, PROJECCITEM.projeccdesc, " _
& "PROJECC.proj_no " _
& "FROM (PROJECC INNER JOIN PROJECCITEM ON PROJECC.projeccid = PROJECCITEM.projeccid) " _
& "INNER JOIN ECC ON PROJECCITEM.ecctabid = ECC.ecctabid " _
& "WHERE (((PROJECC.proj_no)='" & seleproj.Column(1) & "')) " _
& "ORDER BY ECC.citation;"


Set rstProjECC = DB.OpenRecordset(strProjECC)
Debug.Print rn
Do Until rstProjECC.EOF = True
    ListBox2.AddItem rstProjECC!PROJ_NO
    ListBox2.List(ListBox2.ListCount - 1, 1) = rstProjECC!citation
    ListBox2.List(ListBox2.ListCount - 1, 2) = rstProjECC!provision
    ListBox2.List(ListBox2.ListCount - 1, 3) = rstProjECC!projeccdesc
    rstProjECC.MoveNext
Loop
rstProjECC.Close
Set rstProjECC = Nothing
DB.Close
Set DB = Nothing
End Sub

