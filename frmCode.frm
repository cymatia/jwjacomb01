VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCode 
   Caption         =   "FRMCODE"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   OleObjectBlob   =   "frmCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCode"
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

Private Sub CommandButton4_Click()

End Sub

Private Sub CommandButton8_Click()
writeRoom

End Sub

Public Sub UserForm_Initialize()
Dim DB As Database
Dim rstProj As Recordset
Dim strProj As String

' Define the workgroup information file (system database) you will use.
'DBEngine.SystemDB = "C:\Documents and Settings\Jinwoo\Application Data\Microsoft\Access\System6.mdw"
'DBEngine.SystemDB = "C:\Documents and Settings\jwjang\Application Data\Microsoft\Access\system.mdw"
DBEngine.SystemDB = "\\jwja-svr-10\jdb\JW0610.mdw"
' OPTIONAL: Create a workspace with the correct login information
'Set WS = CreateWorkspace("NewWS", "MyID", "MyPwd", dbUseJet)
' OPTIONAL: If you set the workspace object, you can use the new
'           workspace with the OpenDatabase command. For example:
'Set Db = WS.OpenDatabase("C:\My Documents\MyDB.mdb")
'

'Set DB = OpenDatabase("H:\db\db_misc\cl_org_outspec_02.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_outspec\X-CL_ORG_OUTSPEC_02.mdb")

'X-CL_ORG_OUTSPEC_02
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")

'Set Rs = DB.OpenRecordset("project")

'strPCSI = "SELECT PROJSPEC.PROJSPECID, PROJSPEC.PROJ_NO, PROJSPEC.NOTE, " _
& "PROJSPEC.PROJSPECDATE, PROJSPEC.NOTE, Project.PROJ_NAME " _
& "FROM PROJSPEC INNER JOIN Project ON PROJSPEC.PROJ_NO = Project.PROJ_NO " _
& "ORDER BY PROJSPEC.PROJ_NO"

strProj = "SELECT projcode.projcodeid, PROJCODE.PROJ_NO, Project.PROJ_NAME " _
& "FROM Project INNER JOIN PROJCODE ON Project.PROJ_NO = PROJCODE.PROJ_NO " _
& "ORDER BY Project.INPUT DESC"



Set rstProj = DB.OpenRecordset(strProj)

Do Until rstProj.EOF = True
    seleproj.AddItem rstProj!projcodeid
    seleproj.List(seleproj.ListCount - 1, 1) = rstProj!PROJ_NO
    seleproj.List(seleproj.ListCount - 1, 2) = rstProj!PROJ_NAME
    rstProj.MoveNext
Loop
rstProj.Close
Set rstProj = Nothing
DB.Close
Set DB = Nothing

' OPTIONAL: If you set the WS object, use the following to clean up
'WS.Close
'Set WS = Nothing

MsgBox "Project List for Code LOADED!"


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
x = showARDU
writeCode (x)
'frmWorkItem.HIDE
'End If

End Sub



Private Sub seleproj_Change()
'frmCode.Repaint
'CommandButton4_Click 'POPULATE CSI
Debug.Print seleproj
End Sub
