VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWorkItem 
   Caption         =   "frmWorkItem"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7365
   OleObjectBlob   =   "frmWorkItem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWorkItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public wksObj As Workspace
Public DB As Database
Public dbsObj As Database
Public tblObj As TableDef
Public fldObj As Field
Public rstObj, rstProj As Recordset



Private Sub CheckBox1_Click()

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
If mCad = False Then
writeAssWorkItemImg 'DEFAULT
Else
writeAssWorkItemImgCAD
End If
'frmWorkItem.HIDE
'End If

End Sub




Private Sub mCategory_Click()
CommandButton4_Click
End Sub

Public Sub UserForm_Initialize()
'Dim WS As Workspace
Dim DB As Database
Dim rstProj As Recordset
Dim strProj As String

' Define the workgroup information file (system database) you will use.
'DBEngine.SystemDB = "C:\Documents and Settings\Jinwoo\Application Data\Microsoft\Access\System6.mdw"
'DBEngine.SystemDB = "C:\Documents and Settings\jwjang\Application Data\Microsoft\Access\system.mdw"

'DCHANGE

'DBEngine.SystemDB = "H:\DB\JW0610.mdw"
DBEngine.SystemDB = "\\jwja-svr-10\jdb\JW0610.mdw"
' OPTIONAL: Create a workspace with the correct login information
'Set WS = CreateWorkspace("NewWS", "MyID", "MyPwd", dbUseJet)
' OPTIONAL: If you set the workspace object, you can use the new
'           workspace with the OpenDatabase command. For example:
'Set Db = WS.OpenDatabase("C:\My Documents\MyDB.mdb")
'
'Set DB = OpenDatabase("H:\db\db_est\H_H_EST_2013_001.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_est\x_EST_2013_001.mdb")

'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")

'Set Rs = DB.OpenRecordset("project")

'strPCSI = "SELECT PROJSPEC.PROJSPECID, PROJSPEC.PROJ_NO, PROJSPEC.NOTE, " _
& "PROJSPEC.PROJSPECDATE, PROJSPEC.NOTE, Project.PROJ_NAME " _
& "FROM PROJSPEC INNER JOIN Project ON PROJSPEC.PROJ_NO = Project.PROJ_NO " _
& "ORDER BY PROJSPEC.PROJ_NO"

strProj = "SELECT ASS.proj_no, Project.PROJ_NAME " _
& "FROM ASS INNER JOIN Project ON ASS.proj_no = Project.PROJ_NO " _
& "GROUP BY ASS.proj_no, Project.PROJ_NAME, Project.INPUT " _
& "ORDER BY Project.INPUT DESC"



Set rstProj = DB.OpenRecordset(strProj)

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

MsgBox "Project List for WorkItems LOADED!"

End Sub


Private Sub seleproj_Change()
'UserForm1.Repaint
CommandButton4_Click 'POPULATE CSI
End Sub
Public Sub CommandButton4_Click() 'Test Populate CSI
Dim rstProjAss As Recordset
Dim strProjAss As String
Dim DB As Database
Dim rn As Integer
Dim rcArray As Variant
ListBox1.Clear
ListBox1.MultiSelect = fmMultiSelectExtended
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("H:\db\db_misc\JSPEC04_01.mdb")

'Set DB = OpenDatabase("H:\db\db_est\H_H_EST_2013_001.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_est\x_EST_2013_001.mdb")

'strPCSI2 = "SELECT PROJCSI.projcsiid, PROJCSI.PROJ_NO, MCSI.MCSI, MCSI.MCSINAME " _
& "FROM MCSI INNER JOIN PROJCSI ON MCSI.MCSIID = PROJCSI.MCSIID " _
& "WHERE (((PROJCSI.PROJ_NO) = '" & seleproj & "'))" _
& "ORDER BY MCSI.MCSI"
If frmWorkItem.mCategory = True Then
strProjAss = "SELECT ASS.ASSID, ASS.ESTID, ASS.ASSDIV, ASS.ASSNO, ASS.ASSSUBDIV, ASS.ASSNAME, ASS.proj_no, " _
    & "ASS.CATEGORY_ID, CATEGORY.category_no, CATEGORY.category, CATEGORY.category_note " _
    & "FROM ASS INNER JOIN CATEGORY ON ASS.CATEGORY_ID=CATEGORY.category_id " _
& "WHERE (((ASS.proj_no) = '" & seleproj & "'))" _
& "ORDER BY CATEGORY.category_no"
Else
strProjAss = "SELECT ASS.ASSID, ASS.ESTID, ASS.ASSDIV, ASS.ASSNO, ASS.ASSSUBDIV, ASS.ASSNAME, ASS.proj_no " _
& "FROM ASS " _
& "WHERE (((ASS.proj_no) = '" & seleproj & "'))" _
& "ORDER BY ASS.ASSDIV, ASS.ASSNO"

End If

Set rstProjAss = DB.OpenRecordset(strProjAss)
Debug.Print rn
Do Until rstProjAss.EOF = True
    If frmWorkItem.mCategory = True Then
    ListBox1.AddItem rstProjAss!ASSID
    ListBox1.List(ListBox1.ListCount - 1, 1) = rstProjAss!category_id
    ListBox1.List(ListBox1.ListCount - 1, 2) = rstProjAss!category_no
    ListBox1.List(ListBox1.ListCount - 1, 3) = rstProjAss!category
    Else
    ListBox1.AddItem rstProjAss!ASSID
    ListBox1.List(ListBox1.ListCount - 1, 1) = rstProjAss!assdiv
    ListBox1.List(ListBox1.ListCount - 1, 2) = rstProjAss!assno
    ListBox1.List(ListBox1.ListCount - 1, 3) = rstProjAss!ASSNAME
    End If
    rstProjAss.MoveNext
Loop
rstProjAss.Close
Set rstProjAss = Nothing
DB.Close
Set DB = Nothing
End Sub
Private Sub ListBox1_Click()
Dim rstAssItem As Recordset
Dim strAssItem As String
Dim DB As Database
Dim rn As Integer
Dim rcArray As Variant
'Set DB = OpenDatabase("H:\db\db_est\H_H_EST_2013_001.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_est\x_EST_2013_001.mdb")


'Set DB = OpenDatabase("H:\db\db_misc\JSPEC04_01.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")
'strPpart = "SELECT PROJPART.PROJCSIID, PROJPART.MPARTID, PROJPART.PROJPARTNO, MPART.MPART " _
& "FROM PROJPART INNER JOIN MPART ON PROJPART.MPARTID = MPART.MPARTID " _
& "WHERE (((PROJPART.PROJCSIID)=" & ListBox1 & "));"

strAssItem = "SELECT WORK.ASSID, WORK.WKID, WORK.WKCSINO, WORK.WKCSINAME, WORK.WKNOTE " _
& "FROM ASS INNER JOIN [WORK] ON ASS.ASSID = WORK.ASSID " _
& "WHERE (((WORK.ASSID) =" & ListBox1 & "))" _
& "ORDER BY WORK.WKCSINO;"

Set rstAssItem = DB.OpenRecordset(strAssItem)
Debug.Print rstAssItem.RecordCount

ListBox2.Clear

Do Until rstAssItem.EOF = True
    ListBox2.AddItem rstAssItem!wkid
    ListBox2.List(ListBox2.ListCount - 1, 1) = rstAssItem!wkcsino

    rstAssItem.MoveNext
Loop
'RsPart.Close
rstAssItem.Close
Set rstAssItem = Nothing
DB.Close
Set DB = Nothing
End Sub

Private Sub CommandButton6_Click()
Dim x As Integer
'If IsNull(ListBox1) = True Then
'MsgBox "Must select a Assembly Item from the list!"
'Exit Sub
'Else
'x = ListBox1
writeWorkItemImg
'frmWorkItem.HIDE
'End If


End Sub



