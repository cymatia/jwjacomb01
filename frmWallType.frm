VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWallType 
   Caption         =   "UserForm1"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8550
   OleObjectBlob   =   "frmWallType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWallType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








'==========================================================================================================
' 9/21/18
'==========================================================================================================
Public Sub UserForm_Initialize()
'Dim WS As Workspace
Dim DB As Database
Dim rstSeleProj As Recordset
Dim strSeleProj As String

' Define the workgroup information file (system database) you will use.
'DBEngine.SystemDB = "C:\Documents and Settings\Jinwoo\Application Data\Microsoft\Access\System6.mdw"
'DBEngine.SystemDB = "C:\Documents and Settings\jwjang\Application Data\Microsoft\Access\system.mdw"
DBEngine.SystemDB = "\\jwja-svr-10\jdb\JW0610.mdw"
' OPTIONAL: Create a workspace with the correct login information
'Set WS = CreateWorkspace("NewWS", "MyID", "MyPwd", dbUseJet)
' OPTIONAL: If you set the workspace object, you can use the new
'           workspace with the OpenDatabase command. For example:
'Set Db = WS.OpenDatabase("C:\My Documents\MyDB.mdb")

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

Set DB = OpenDatabase("\\jwja-svr-10\jdb\DB_WALL\WALLASS03.mdb")
'Set DB = OpenDatabase("H:\db\db_misc\JSPEC04_01.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

'Set Rs = DB.OpenRecordset("project")

strSeleProj = "SELECT ProjBldgEnv.ProjBldgEnvID, ProjBldgEnv.Proj_no, Project.PROJ_NAME " _
& "FROM Project INNER JOIN ProjBldgEnv ON Project.PROJ_NO = ProjBldgEnv.Proj_no " _
& "ORDER BY ProjBldgEnv.proj_no DESC"

Set rstSeleProj = DB.OpenRecordset(strSeleProj)

Do Until rstSeleProj.EOF = True

    seleproj.AddItem rstSeleProj!ProjBldgEnvid
    
    seleproj.List(seleproj.ListCount - 1, 1) = rstSeleProj!PROJ_NO
    seleproj.List(seleproj.ListCount - 1, 2) = rstSeleProj!PROJ_NAME
    rstSeleProj.MoveNext
Loop
rstSeleProj.Close



Set rstSeleProj = Nothing
DB.Close
Set DB = Nothing

' OPTIONAL: If you set the WS object, use the following to clean up
'WS.Close
'Set WS = Nothing

MsgBox "LOADED!"
End Sub

Private Sub seleproj_Change()
'==========================================================================================================
' 9/21/18
'==========================================================================================================

Me.Repaint
Dim DB As Database
Dim rstProjBldgEnvItem  As Recordset
Dim strProjBldgEnvItem  As String
Dim rn As Integer

DBEngine.SystemDB = "\\jwja-svr-10\jdb\JW0610.mdw"
Set DB = OpenDatabase("\\jwja-svr-10\jdb\DB_WALL\WALLASS03.mdb")
'=============================================================================================================
' Show all ProjBldgEnv
'=============================================================================================================
strProjBldgEnvItem = "SELECT ProjBldgEnv.Proj_no, ProjBldgEnv.ProjBldgEnvID, ProjBldgEnv.ProjBldgEnvName " _
& "FROM ProjBldgEnv " _
& "WHERE (((ProjBldgEnv.ProjBldgEnvID)=" & seleproj & "))"
'=============================================================================================================
'=============================================================================================================


Me.lsProjBldgEnv.Clear
Set rstProjBldgEnvItem = DB.OpenRecordset(strProjBldgEnvItem)
'Debug.Print rn
Do Until rstProjBldgEnvItem.EOF = True
    lsProjBldgEnv.AddItem rstProjBldgEnvItem!ProjBldgEnvid
    'istseleCode.List(listseleCode.ListCount - 1, 1) = rstProjCodeItem!vCahpterNo
    lsProjBldgEnv.List(lsProjBldgEnv.ListCount - 1, 1) = rstProjBldgEnvItem!ProjBldgEnvName
    rstProjBldgEnvItem.MoveNext
Loop

rstProjBldgEnvItem.Close

Set rstProjBldgEnvItem = Nothing

DB.Close
Set DB = Nothing
End Sub

Private Sub lsProjBldgEnv_Click()
Me.Repaint
Dim DB As Database
Dim rstProjEnvItem As Recordset
Dim strProjEnvItem As String
Dim rn As Integer

DBEngine.SystemDB = "\\jwja-svr-10\jdb\JW0610.mdw"
Set DB = OpenDatabase("\\jwja-svr-10\jdb\DB_WALL\WALLASS03.mdb")

'=============================================================================================================
'=============================================================================================================
' Show ProjEnvItems
'=============================================================================================================
'
strProjEnvItem = "SELECT ConnProjBldgEnv.ProjBldgEnvid, ConnProjBldgEnv.ProjENvID, ProjEnvItem.ProjEnvNo, " _
& "ProjEnvItem.ProjENvCategory, ProjEnvItem.ProjEnvType, ProjEnvItem.ProjENvName " _
& "FROM ProjEnvItem INNER JOIN ConnProjBldgEnv ON ProjEnvItem.ProjEnvID = ConnProjBldgEnv.ProjENvID " _
& "WHERE (((ConnProjBldgEnv.ProjBldgEnvid)=" & lsProjBldgEnv & "));"
Me.lsEnvItem.Clear
Set rstProjEnvItem = DB.OpenRecordset(strProjEnvItem)

Do Until rstProjEnvItem.EOF = True
    lsEnvItem.AddItem rstProjEnvItem!ProjEnvID
    'istseleCode.List(listseleCode.ListCount - 1, 1) = rstProjCodeItem!vCahpterNo
    If IsNull(rstProjEnvItem!ProjEnvName) = False Then
    
    lsEnvItem.List(lsEnvItem.ListCount - 1, 1) = rstProjEnvItem!ProjEnvName
    Else
    lsEnvItem.List(lsEnvItem.ListCount - 1, 1) = ""
    End If
    rstProjEnvItem.MoveNext
Loop
rstProjEnvItem.Close
Set rstProjEnvItem = Nothing
DB.Close
Set DB = Nothing

End Sub
Private Sub cmdWritheWallType_Click()
Dim x, y, mProjBldgEnvID As Integer
If IsNull(lsProjBldgEnv) = True Then
    MsgBox "Must select a Proj Bldg Env Item from the list!"
    Exit Sub
Else
    'x = lsProjBldgEnv
    mProjBldgEnvID = lsProjBldgEnv
    y = writeProjENV(mProjBldgEnvID) '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    frmWallType.Hide
End If
End Sub

