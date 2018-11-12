VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProjCode 
   Caption         =   "frmProjCode"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   OleObjectBlob   =   "frmProjCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProjCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================================================================

Private Sub cmdWriteProjCode_Click()
Dim x As Integer
If IsNull(listseleCode) = True Then
    MsgBox "Must select a Chapter from the list!"
    Exit Sub
Else
    x = listseleCode
    mprojcodeID = seleproj
    y = writeProjCode(mprojcodeID, x) '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    frmProjCode.Hide
End If

End Sub

Private Sub cmdWriteProjCodeNW_Click()
Dim x As Integer
If IsNull(listseleCode) = True Then
    MsgBox "Must select a Chapter from the list!"
    Exit Sub
Else
    x = listseleCode
    mprojcodeID = seleproj
    y = writeProjCodeNW(mprojcodeID, x) '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    frmProjCode.Hide
End If
End Sub


Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub seleproj_Change()
Me.Repaint
Dim DB As Database
Dim rstProjCodeItem  As Recordset
Dim strProjCodeItem As String
Dim rn As Integer

DBEngine.SystemDB = "\\jwja-svr-10\jdb\JW0610.mdw"
Set DB = OpenDatabase("\\jwja-svr-10\jdb\DB_CODE\2018-Code-01.mdb")
'=============================================================================================================
' Show all laws
'=============================================================================================================
strProjCodeItem_1 = "SELECT rellaw.relLawId, Val([chapterno]) AS vCahpterNo, rellaw.lawno, rellaw.projcodeid " _
& "FROM rellaw " _
& "WHERE (((rellaw.projcodeid) =" & seleproj & ")) " _
& "ORDER BY Val([chapterno]), rellaw.lawno;"

'=============================================================================================================
'=============================================================================================================
' Show Chapters
'=============================================================================================================
strProjCodeItem = "SELECT rellaw.chapterid, Val([chapterno]) AS vCahpterNo, DLookUp('chaptername','chapter','chapterid=' & [chapterid]) AS chaptername, " _
& "rellaw.projcodeid " _
& "FROM rellaw " _
& "GROUP BY rellaw.chapterid, Val([chapterno]), DLookUp('chaptername','chapter','chapterid=' & [chapterid]), rellaw.projcodeid " _
& "HAVING (((rellaw.projcodeid) = " & seleproj & ")) " _
& "ORDER BY Val([chapterno]);"

'=============================================================================================================
'=============================================================================================================


Me.listseleCode.Clear
Set rstProjCodeItem = DB.OpenRecordset(strProjCodeItem)
Debug.Print rn
Do Until rstProjCodeItem.EOF = True
    listseleCode.AddItem rstProjCodeItem!chapterid
    listseleCode.List(listseleCode.ListCount - 1, 1) = rstProjCodeItem!vCahpterNo
    listseleCode.List(listseleCode.ListCount - 1, 2) = rstProjCodeItem!chaptername
    rstProjCodeItem.MoveNext
Loop

End Sub

'==========================================================================================================
' 9/21/18
'==========================================================================================================
Public Sub UserForm_Initialize()
'Dim WS As Workspace
Dim DB As Database
Dim RstProjCode, rstSeleProj As Recordset
Dim strProjCode, strSeleProj As String

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

Set DB = OpenDatabase("\\jwja-svr-10\jdb\DB_CODE\2018-Code-01.mdb")
'Set DB = OpenDatabase("H:\db\db_misc\JSPEC04_01.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
' 18/11/11
'Set Rs = DB.OpenRecordset("project")

strSeleProj = "SELECT projcode.projcodeid, projcode.projcodename, projcode.proj_no, Project.PROJ_NAME " _
& "FROM projcode INNER JOIN Project ON projcode.proj_no = Project.PROJ_NO " _
& "ORDER BY projcode.proj_no DESC;"

Set rstSeleProj = DB.OpenRecordset(strSeleProj)

Do Until rstSeleProj.EOF = True

    seleproj.AddItem rstSeleProj!projcodeid
    If IsNull(rstSeleProj!projcodename) = True Then
    seleproj.List(seleproj.ListCount - 1, 1) = ""
    Else
    seleproj.List(seleproj.ListCount - 1, 1) = rstSeleProj!projcodename
    End If
    seleproj.List(seleproj.ListCount - 1, 2) = rstSeleProj!PROJ_NO
    seleproj.List(seleproj.ListCount - 1, 3) = rstSeleProj!PROJ_NAME
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
