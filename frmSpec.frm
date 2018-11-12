VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSpec 
   Caption         =   "UserForm1"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   OleObjectBlob   =   "frmSpec.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSpec"
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


Private Sub CommandButton1_Click() ' TEST
Dim x As Integer

x = ListBox1
y = printspec(x)
End Sub

Public Sub CommandButton4_Click() 'Test Populate CSI
Dim RsPCSI2 As Recordset
Dim strPCSI2 As String
Dim DB As Database
Dim rn As Integer
Dim rcArray As Variant
ListBox1.Clear
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_SPEC\X_JSPEC_17_01.mdb")
strPCSI2 = "SELECT PROJCSI.projcsiid, PROJCSI.PROJ_NO, MCSI.MCSI, MCSI.MCSINAME " _
& "FROM MCSI INNER JOIN PROJCSI ON MCSI.MCSIID = PROJCSI.MCSIID " _
& "WHERE (((PROJCSI.PROJ_NO) = '" & seleproj & "'))" _
& "ORDER BY MCSI.MCSI"
Set RsPCSI2 = DB.OpenRecordset(strPCSI2)
Debug.Print rn
Do Until RsPCSI2.EOF = True
    ListBox1.AddItem RsPCSI2!PROJCSIID
    ListBox1.List(ListBox1.ListCount - 1, 1) = RsPCSI2!mCSI
    ListBox1.List(ListBox1.ListCount - 1, 2) = RsPCSI2!MCSIname
    RsPCSI2.MoveNext
Loop
End Sub

Private Sub CommandButton5_Click() ' INSERT TEXT
Dim returnPnt As Variant
Dim width As Double
Dim text As String
Debug.Print ListBox1
 UserForm1.Hide
returnPnt = ThisDrawing.Utility.GetPoint(, "Enter a point: ")

    'corner(0) = 0#: corner(1) = 10#: corner(2) = 0#
    width = 10
    text = TextBox1
    ' Creates the mtext Object
    Set mtextObj = ThisDrawing.ModelSpace.AddMText(returnPnt, width, text)
    'ZoomAll
    ThisDrawing.Regen acActiveViewport
    UserForm1.Hide
End Sub


'**********************************************************************************************************
'**********************************************************************************************************
'
'12/22/17
'
'**********************************************************************************************************
'**********************************************************************************************************

Private Sub CommandButton6_Click() 'WRIET MTEXT
Dim x As Integer
If IsNull(ListBox1) = True Then
    MsgBox "Must select a CSI from the list!"
    Exit Sub
Else
    x = ListBox1
    y = writemtext(x) '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    frmSpec.Hide
End If

End Sub





Public Sub ListBox1_Click()
Dim RsPart As Recordset
Dim strPpart As String
Dim DB As Database
Dim rn As Integer
Dim rcArray As Variant

'**********************************************************************************************************
'**********************************************************************************************************

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_SPEC\X_JSPEC_17_01.mdb")
'Set DB = OpenDatabase("H:\db\db_misc\JSPEC04_01.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")

'**********************************************************************************************************
'**********************************************************************************************************

strPpart = "SELECT PROJPART.PROJCSIID, PROJPART.MPARTID, PROJPART.PROJPARTNO, MPART.MPART " _
& "FROM PROJPART INNER JOIN MPART ON PROJPART.MPARTID = MPART.MPARTID " _
& "WHERE (((PROJPART.PROJCSIID)=" & ListBox1 & "));"
Set RsPart = DB.OpenRecordset(strPpart)
Debug.Print RsPart.RecordCount

ListBox2.Clear

Do Until RsPart.EOF = True
    ListBox2.AddItem RsPart!MPARTID
    ListBox2.List(ListBox2.ListCount - 1, 1) = RsPart!MPART

    RsPart.MoveNext
Loop
'RsPart.Close

End Sub
Sub selectpt()
Dim returnPnt As Variant
Dim width As Double
Dim text As String
Debug.Print ListBox1
returnPnt = ThisDrawing.Utility.GetPoint(, "Enter a point: ")

    'corner(0) = 0#: corner(1) = 10#: corner(2) = 0#
    width = 10
    text = ListBox1
    ' Creates the mtext Object
    Set mtextObj = ThisDrawing.ModelSpace.AddMText(returnPnt, width, text)
    ZoomAll
ThisDrawing.Regen acActiveViewport
End Sub


'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
'
'12/22/17 NOT NEEDED
'
'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ

Sub opendb()
Dim WS As Workspace
Dim DB As Database
Dim rs As Recordset
Dim strPCSI As String

'Define the workgroup information file (system database) you will use.

'DBEngine.SystemDB = "H:\DB\DB_MISC\SECURITY.mdw"
'DBEngine.SystemDB = "C:\Documents and Settings\Jinwoo\Application Data\Microsoft\Access\System6.mdw"
'DBEngine.SystemDB = "C:\Documents and Settings\jwjang\Application Data\Microsoft\Access\System.mdw"
'DBEngine.SystemDB = "h:\DB\JW0610.mdw"
DBEngine.SystemDB = "\\jwja-svr-10\jdb\DB\JW0610.mdw"

'OPTIONAL: Create a workspace with the correct login information
'Set WS = CreateWorkspace("NewWS", "MyID", "MyPwd", dbUseJet)
'OPTIONAL: If you set the workspace object, you can use the new
'workspace with the OpenDatabase command. For example:

'Set Db = WS.OpenDatabase("C:\My Documents\MyDB.mdb")
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_SPEC\X_JSPEC_17_01.mdb")
'Set DB = OpenDatabase("H:\db\db_SPEC\H_H_JSPEC04_02.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")
'Set Rs = DB.OpenRecordset("project")

strPCSI = "SELECT PROJSPEC.PROJSPECID, PROJSPEC.PROJ_NO, PROJSPEC.NOTE, " _
 & "PROJSPEC.PROJSPECDATE, PROJSPEC.NOTE, Project.PROJ_NAME " _
 & "FROM PROJSPEC INNER JOIN Project ON PROJSPEC.PROJ_NO = Project.PROJ_NO " _
 & "ORDER BY PROJSPEC.PROJSPECID DESC"
 
Set rs = DB.OpenRecordset(strPCSI)
Do Until rs.EOF = True
    seleproj.AddItem rs!PROJ_NO
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
DB.Close
Set DB = Nothing
'OPTIONAL: If you set the WS object, use the following to clean up
'WS.Close
'Set WS = Nothing
MsgBox "Done"
End Sub
'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ

'**********************************************************************************************************
'**********************************************************************************************************
'
'12/22/17
'
'**********************************************************************************************************
'**********************************************************************************************************

Private Sub ListBox2_Click()
Dim RsPart As Recordset
Dim strPpart As String
Dim DB As Database
Dim rn As Integer
Dim rcArray As Variant

'**********************************************************************************************************
'**********************************************************************************************************
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_SPEC\X_JSPEC_17_01.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")
'**********************************************************************************************************
'**********************************************************************************************************

strPpart = "SELECT PROJPART.PROJCSIID, PROJPART.MPARTID, PROJPART.PROJPARTNO, MPART.MPART " _
& "FROM PROJPART INNER JOIN MPART ON PROJPART.MPARTID = MPART.MPARTID " _
& "WHERE (((PROJPART.PROJCSIID)=" & ListBox1 & "));"
Set RsPart = DB.OpenRecordset(strPpart)
Debug.Print RsPart.RecordCount

'TextBox1.Clear

Do Until RsPart.EOF = True
    TextBox1 = RsPart!MPART
    RsPart.MoveNext
Loop
RsPart.Close

End Sub

Private Sub seleproj_Change()
frmSpec.Repaint
CommandButton4_Click 'POPULATE CSI

End Sub

'**********************************************************************************************************
'**********************************************************************************************************
'
'12/22/17
'
'**********************************************************************************************************
'**********************************************************************************************************


Public Sub UserForm_Initialize()
'Dim WS As Workspace
Dim DB As Database
Dim RsPCSI As Recordset
Dim strPCSI As String

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

Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_SPEC\X_JSPEC_17_01.mdb")
'Set DB = OpenDatabase("H:\db\db_misc\JSPEC04_01.mdb")
'Set DB = OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
'Set DB = OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

'Set Rs = DB.OpenRecordset("project")

strPCSI = "SELECT PROJSPEC.PROJSPECID, PROJSPEC.PROJ_NO, PROJSPEC.NOTE, " _
& "PROJSPEC.PROJSPECDATE, PROJSPEC.NOTE, Project.PROJ_NAME " _
& "FROM PROJSPEC INNER JOIN Project ON PROJSPEC.PROJ_NO = Project.PROJ_NO " _
& "ORDER BY PROJSPEC.PROJ_NO desc"
Set RsPCSI = DB.OpenRecordset(strPCSI)

Do Until RsPCSI.EOF = True
    seleproj.AddItem RsPCSI!PROJ_NO
    seleproj.List(seleproj.ListCount - 1, 1) = RsPCSI!PROJ_NAME
    RsPCSI.MoveNext
Loop
RsPCSI.Close
Set RsPCSI = Nothing
DB.Close
Set DB = Nothing

' OPTIONAL: If you set the WS object, use the following to clean up
'WS.Close
'Set WS = Nothing

MsgBox "LOADED!"

End Sub



Sub ComboBox1_afterupdate()
strPCSI2 = "SELECT PROJCSI.PROJ_NO, MCSI.MCSI, MCSI.MCSINAME " _
& "FROM MCSI INNER JOIN PROJCSI ON MCSI.MCSIID = PROJCSI.MCSIID " _
& " ON PROJSPEC.PROJ_NO = Project.PROJ_NO" _
 & "WHERE (((PROJCSI.PROJ_NO) = " & seleproj & ")) " _
 & "ORDER BY MCSI.MCSI"
Set RsPCSI2 = DB.OpenRecordset(strPCSI2)

Do Until RsPCSI2.EOF = True
    ListBox1.AddItem RsPCSI2!PROJ_NO
    RsPCSI2.MoveNext
Loop
End Sub



Private Sub CommandButton2_Click()
 On Error Resume Next
   
    Set dbsObj = DBEngine.Workspaces(0).OpenDatabase("\\jwja-svr-10\jdb\db_SPEC\X_JSPEC_17_01.mdb")
  
   'Set dbsObj = DBEngine.Workspaces(0).OpenDatabase("X:\db_misc\02_Jspec_B_001.mdb")
      'Set dbsObj = DBEngine.Workspaces(0).OpenDatabase("h:\db\db_misc\JSPEC0310.mdb")

   Set rstProj = dbsObj.OpenRecordset("project", dbOpenDynaset)

   strPCSI = "SELECT PROJSPEC.PROJSPECID, PROJSPEC.PROJ_NO, PROJSPEC.NOTE, " _
    & "PROJSPEC.PROJSPECDATE, PROJSPEC.NOTE, Project.PROJ_NAME " _
    & "FROM PROJSPEC INNER JOIN Project ON PROJSPEC.PROJ_NO = Project.PROJ_NO"
    Set rstPCSI = dbsObj.OpenRecordset(strPCSI)

   Set dbsObj = Nothing
End Sub

Private Sub CommandButton3_Click() 'QUIT
Unload Me
End Sub

