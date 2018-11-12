VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDoor 
   Caption         =   "Door Schedule Insert"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   OleObjectBlob   =   "frmDoor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDoor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public WS As Workspace
Public DB As Database
Public dbsObj As Database
Public tblObj As TableDef
Public fldObj As Field
Public rstObj, rstProj As Recordset


Private Sub CommandButton8_Click()
writeRoom
End Sub



Public Sub UserForm_Initialize()
Dim WS As Workspace
Dim DB As Database
Dim RsPDR, RsPPT As Recordset
Dim strPDR, strPPT As String

' Define the workgroup information file (system database) you will use.
'DBEngine.SystemDB = "C:\Documents and Settings\Jinwoo\Application Data\Microsoft\Access\System6.mdw"
'DBEngine.SystemDB = "C:\Documents and Settings\jwjang\Application Data\Microsoft\Access\SECURED.mdw"
'DBEngine.SystemDB = "C:\Documents and Settings\JINWOO\Application Data\Microsoft\Access\SECURED.mdw"
'DBEngine.SystemDB = "C:\Documents and Settings\jwjang\Application Data\Microsoft\Access\system.mdw"
' 062010 AT UPSTAIR SYSTEM7.MDW
'DBEngine.SystemDB = "H:\DB\JW0610.mdw"

DBEngine.SystemDB = "\\jwja-svr-10\jdb\JW0610.mdw"

' OPTIONAL: Create a workspace with the correct login information
'Set WS = CreateWorkspace("NewWS", "MyID", "MyPwd", dbUseJet)
' OPTIONAL: If you set the workspace object, you can use the new
'           workspace with the OpenDatabase command. For example:
'Set Db = WS.OpenDatabase("C:\My Documents\MyDB.mdb")

'**********************************************************
'Set DB = OpenDatabase("h:\db\db_misc\Door02_01.mdb")
'**********************************************************
'**********************************************************
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\Door02_01.mdb")
'**********************************************************


'Set DB = OpenDatabase("h:\db\DATAS\DT11_JPM_01.mdb")

'Set Rs = DB.OpenRecordset("project")

'strPDR = "SELECT Project.PROJ_NO, Project.PROJ_NAME, Project.STATUS, Project.INPUT " _
& "FROM DRGP INNER JOIN Project ON DRGP.PROJ_NO = Project.PROJ_NO " _
& "GROUP BY Project.PROJ_NO, Project.PROJ_NAME, Project.STATUS, Project.INPUT " _
& "HAVING (((Project.Status) = 'act'))" _
& "ORDER BY Project.INPUT DESC;"

strPDR = "SELECT Project.PROJ_NO, Project.PROJ_NAME, Project.STATUS, Project.INPUT " _
& "FROM Project  " _
& "GROUP BY Project.PROJ_NO, Project.PROJ_NAME, Project.STATUS, Project.INPUT " _
& "HAVING (((Project.Status) = 'act'))" _
& "ORDER BY Project.INPUT DESC;"

Set RsPDR = DB.OpenRecordset(strPDR)
Do Until RsPDR.EOF = True
    seleproj.AddItem RsPDR!PROJ_NO
    seleproj.List(seleproj.ListCount - 1, 1) = RsPDR!PROJ_NAME
    RsPDR.MoveNext
Loop
RsPDR.Close
Set RsPPT = Nothing
DB.Close
Set DB = Nothing
'**********************************************************
'**********************************************************
'**********************************************************

'**********************************************************
'Set DB = OpenDatabase("h:\db\db_misc\Partition.mdb")
'**********************************************************

'**********************************************************
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\Partition.mdb")
'**********************************************************



strPPT = "SELECT Project.PROJ_NO, Project.PROJ_NAME, Project.STATUS, Project.INPUT " _
& "FROM ProPart INNER JOIN Project ON ProPart.PROJ_NO = Project.PROJ_NO " _
& "GROUP BY Project.PROJ_NO, Project.PROJ_NAME, Project.STATUS, Project.INPUT " _
& "HAVING (((Project.Status) = 'act'))" _
& "ORDER BY Project.INPUT DESC;"
Set RsPPT = DB.OpenRecordset(strPPT)
Do Until RsPPT.EOF = True
    seleprojpart.AddItem RsPPT!PROJ_NO
    seleprojpart.List(seleprojpart.ListCount - 1, 1) = RsPPT!PROJ_NAME
    RsPPT.MoveNext
Loop
RsPPT.Close
Set RsPPT = Nothing
DB.Close
Set DB = Nothing

'**********************************************************
'**********************************************************
'**********************************************************



' OPTIONAL: If you set the WS object, use the following to clean up
'WS.Close
'Set WS = Nothing

MsgBox "Projects Loaded!"
frmDoor.Hide
End Sub
Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub CommandButton2_Click()
mkDr
End Sub

Private Sub Example()
' Define the block
Dim blockObj As AcadBlock
Dim insertionPnt(0 To 2) As Double
insertionPnt(0) = 0
insertionPnt(1) = 0
insertionPnt(2) = 0
Set blockObj = ThisDrawing.Blocks.Add _
          (insertionPnt, "BlockWithAttribute")

' Add an attribute to the block
Dim attributeObj As AcadAttribute
Dim height As Double
Dim mode As Long
Dim prompt As String
Dim insertionPoint(0 To 2) As Double
Dim tag As String
Dim value As String
height = 1
mode = acAttributeModeVerify
prompt = "New Prompt"
insertionPoint(0) = 5
insertionPoint(1) = 5
insertionPoint(2) = 0
tag = "New Tag"
value = "New Value"
Set attributeObj = blockObj.AddAttribute(height, mode, _
               prompt, insertionPoint, tag, value)
' Insert the block, creating a block reference
' and an attribute reference
Dim blockRefObj As AcadBlockReference
insertionPnt(0) = 2
insertionPnt(1) = 2
insertionPnt(2) = 0
Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
    (insertionPnt, "BlockWithAttribute", 1#, 1#, 1#, 0)

End Sub


Private Sub CommandButton3_Click()
MkHDWRimg
End Sub

Private Sub CommandButton4_Click()
Partition
End Sub




Private Sub CommandButton5_Click()
If Me.writeDoor = True Then
    mkDr
ElseIf Me.writeWindow = True Then
    writeWinImg
End If

End Sub

Private Sub ListBox2_Click()
Dim DB As Database
Dim RsProjWinItem As Recordset
Dim strProjWinItem As String
Dim rn As Integer
Dim rcArray As Variant
ListBox1.Clear

'**********************************************************
'Set DB = OpenDatabase("h:\db\db_misc\Door02_01.mdb")
'**********************************************************
'**********************************************************
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\Door02_01.mdb")
'**********************************************************



strProjWinItem = "SELECT projwin.projwinid, connprojwin.connprojwinid, connprojwin.projwinno, " _
& "WINDOW.WINSIZE, WINDOW.CATNO, WINDOW.WINMAT, WINDOW.WINTYPE, WINDOW.WINMODEL, WINDOW.WINAREA, " _
& "WINDOW.WINOPENAREA, WINDOW.WINIMGNAME " _
& "FROM (projwin INNER JOIN connprojwin ON projwin.projwinid = connprojwin.projwinid) " _
& "INNER JOIN WINDOW ON connprojwin.winid = WINDOW.WINID " _
& "WHERE (((projwin.projwinid)=" & ListBox2 & ")) " _
& "ORDER BY connprojwin.PROJWINNO;"
Set RsProjWinItem = DB.OpenRecordset(strProjWinItem)
rn = RsProjWinItem.RecordCount
Debug.Print rn
Do Until RsProjWinItem.EOF = True
    ListBox1.AddItem RsProjWinItem!projwinid
    ListBox1.List(ListBox1.ListCount - 1, 1) = RsProjWinItem!projwinid
    'ListBox1.List(ListBox1.ListCount - 1, 2) = RsDRG!DRGNOTE
    RsProjWinItem.MoveNext
Loop

End Sub

Private Sub seleproj_Change()
frmDoor.Repaint
Dim DB As Database
Dim RsDRG As Recordset
Dim strDRG As String
Dim rn As Integer
Dim rcArray As Variant
ListBox1.Clear
'**********************************************************
'Set DB = OpenDatabase("h:\db\db_misc\Door02_01.mdb")
'**********************************************************
'**********************************************************
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\Door02_01.mdb")
'**********************************************************


If Me.writeDoor = True Then

    strDRG = "SELECT DRGP.DRGID, DRGP.DRID, DRGP.PROJ_NO, DRGP.DRGNO, DRGP.DRTYPE, " _
    & "DRGP.DRFIN, DRGP.DRFRAME, DRGP.DRJAMB, DRGP.DRHEAD, DRGP.HDSET, DRGP.DRGNOTE, " _
    & "DR.DRWIDTH, DR.DRHEIGHT, DR.DRMAT, DR.DRTHK, DR.DRMFGR, DR.DRFIRE, DR.DRNOTE, DR.OLDSIZE " _
    & "FROM DRGP INNER JOIN DR ON DRGP.DRID = DR.DRID " _
    & "WHERE (((DRGP.PROJ_NO) = '" & seleproj & "')) " _
    & "ORDER BY DRGP.DRGNO;"
    Set RsDRG = DB.OpenRecordset(strDRG)
    Debug.Print rn
    Do Until RsDRG.EOF = True
        ListBox1.AddItem RsDRG!DRGNO
        ListBox1.List(ListBox1.ListCount - 1, 1) = RsDRG!OLDSIZE
        'ListBox1.List(ListBox1.ListCount - 1, 2) = RsDRG!DRGNOTE
        RsDRG.MoveNext
    Loop
ElseIf Me.writeWindow = True Then
    strProjWin = "SELECT projwin.projwinid, projwin.proj_no, projwin.connprojwin, projwin.note, projwin.input " _
    & "FROM projwin " _
    & "WHERE (((projwin.proj_no) = '" & seleproj & "')) " _
    & "ORDER BY projwin.projwinid;"
    Set RsProjWin = DB.OpenRecordset(strProjWin)
    Debug.Print rn
    Do Until RsProjWin.EOF = True
        ListBox2.AddItem RsProjWin!projwinid
        ListBox2.List(ListBox2.ListCount - 1, 1) = RsProjWin!PROJ_NO
        'ListBox1.List(ListBox1.ListCount - 1, 2) = RsDRG!DRGNOTE
        RsProjWin.MoveNext
    Loop

End If

End Sub
Sub Partition()

Dim DB As Database
Dim RstParti As Recordset
Dim strParti, txtRow As String
Dim rnParti As Integer   'ProjSet no.
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
Dim tblWidth As Double
Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
Dim height As Double
Dim mode As Long
Dim prompt, mKey As String
Dim blockRefObj As AcadBlockReference
Dim BlkObj As Object
Dim i, iat As Integer

'frmDoor.ListBox1.Clear
'**********************************************************
'Set DB = OpenDatabase("h:\db\db_misc\partition.mdb")
'**********************************************************

'**********************************************************
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\partition.mdb")
'**********************************************************


strParti = "SELECT PROPART.PRPTID, PROPART.KEY, PROPART.EXTENT, PROPART.FIGURE, " _
& "PROPART.NOTE, PARTITION.THK, PARTITION.FIRE, PARTITION.INSUL, PARTITION.STC, PARTITION.TestNo, " _
& "PARTITION.PTNOTE, Sub.SubDepth, Sub.SubTYPE, CONNLayer.PTID, CONNLayer.SIDE, Layer.LayerTHK, " _
& "Layer.LayerTYPE, PROPART.PROJ_NO " _
& "FROM Sub INNER JOIN ((PARTITION INNER JOIN (Layer INNER JOIN CONNLayer ON Layer.LayerID = CONNLayer.LayerID) " _
& "ON PARTITION.PTID = CONNLayer.PTID) INNER JOIN PROPART ON PARTITION.PTID = PROPART.PTID) ON Sub.SubID = PARTITION.SubID " _
& "GROUP BY PROPART.PRPTID, PROPART.KEY, PROPART.EXTENT, PROPART.FIGURE, PROPART.NOTE, PARTITION.THK, PARTITION.FIRE, " _
& "PARTITION.INSUL, PARTITION.STC, PARTITION.TestNo, PARTITION.PTNOTE, Sub.SubDepth, Sub.SubTYPE, " _
& "CONNLayer.PTID, CONNLayer.SIDE, Layer.LayerTHK, Layer.LayerTYPE, PROPART.PROJ_NO " _
& "HAVING (((PROPART.PROJ_NO)='" & seleprojpart & "')) " _
& "ORDER BY PROPART.KEY;"

RecordSets:
Set RstParti = DB.OpenRecordset(strParti, dbOpenDynaset)

If RstParti.RecordCount = 0 Then
    GoTo ExitPartition
End If

TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify
'InsertJBlock "e:\drawfile\jvba\XHDWRHEAD01.dwg", tblPT, "Model"

'**********************************************************************************************
'ThisDrawing.ModelSpace.InsertBlock tblPT, "H:\drawfile\jvba\XPARTIHEAD01.dwg", 1, 1, 1, 0
'**********************************************************************************************

'**********************************************************************************************
ThisDrawing.ModelSpace.InsertBlock TblPT, "\\jwja-svr-10\drawfile\jvba\XPARTIHEAD01.dwg", 1, 1, 1, 0
'**********************************************************************************************
'**********************************************************************************************
'InsertJBlock "H:\drawfile\jvba\XPARTIROW01.dwg", tblPT2, "Model"
'**********************************************************************************************
'**********************************************************************************************
InsertJBlock "\\jwja-svr-10\drawfile\jvba\XPARTIROW01.dwg", tblPT2, "Model"
'**********************************************************************************************



TblPT(1) = TblPT(1) - 48
RstParti.MoveLast
RstParti.MoveFirst
rnParti = RstParti.RecordCount
Debug.Print rnParti

With RstParti
    For i = 1 To rnParti
    If !Key = mKey Then
    GoTo PARTINEXT
    End If
         Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                     (TblPT, "XPARTIROW01", 1, 1, 1, 0)
        '******************************
          'Get the block's attributes
                Dim varAttributes As Variant
                varAttributes = blockRefObj.GetAttributes
                    For iat = LBound(varAttributes) To UBound(varAttributes)
                    Debug.Print varAttributes(iat).TagString
                        If varAttributes(iat).TagString = "XKEY" Then
                            If IsNull(!Key) = True Then
                                varAttributes(iat).textString = ""
                            Else
                                varAttributes(iat).textString = !Key
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If
                        If varAttributes(iat).TagString = "XDEPTH" Then
                            If IsNull(!THK) = True Then
                                varAttributes(iat).textString = ""
                            Else
                                varAttributes(iat).textString = Num2Frac(!THK) & """"
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If
                        If varAttributes(iat).TagString = "XSUBDEPTH" Then
                            If IsNull(!SUBDEPTH) = True Then
                                varAttributes(iat).textString = ""
                            Else
                                varAttributes(iat).textString = Num2Frac(!SUBDEPTH) & """"
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If
                        If varAttributes(iat).TagString = "XSUBTYPE" Then
                            If IsNull(!SUBTYPE) = True Then
                                varAttributes(iat).textString = ""
                            Else
                                varAttributes(iat).textString = !SUBTYPE
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If
                        If varAttributes(iat).TagString = "XLAYERTHK" Then
                            If IsNull(!LAYERTHK) = True Then
                                varAttributes(iat).textString = ""
                            Else
                                varAttributes(iat).textString = Num2Frac(!LAYERTHK) & """"
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If
                        If varAttributes(iat).TagString = "XFIRE" Then
                            If IsNull(!FIRE) = True Then
                                varAttributes(iat).textString = ""
                            Else
                                varAttributes(iat).textString = !FIRE
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If
                        If varAttributes(iat).TagString = "XINSUL" Then
                            If IsNull(!INSUL) = True Then
                                varAttributes(iat).textString = ""
                            Else
                                varAttributes(iat).textString = !INSUL
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If
                        If varAttributes(iat).TagString = "XSTC" Then
                            If IsNull(!STC) = True Then
                                varAttributes(iat).textString = ""
                            Else
                                varAttributes(iat).textString = !STC
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If
                        If varAttributes(iat).TagString = "XTESTNO" Then
                            If IsNull(!TestNo) = True Then
                                varAttributes(iat).textString = ""
                            Else
                             '********** pre 2014 **********
                             '   If Len(!testno) > 7 Then
                             '   varAttributes(iat).TextString = Left(!testno, 6)
                                
                             '   TestNo2 = Mid(!testno, 7, Len(!testno))
                             '   Else
                             '   varAttributes(iat).TextString = !testno
                             '   End If
                             varAttributes(iat).textString = SEPTEXT2(!TestNo, 1, 2, 6)
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If
                        If varAttributes(iat).TagString = "XTESTNO2" Then
                            If IsNull(!TestNo) = True Then
                                varAttributes(iat).textString = ""
                            Else
                                'varAttributes(iat).TextString = Trim(TestNo2)
                                varAttributes(iat).textString = SEPTEXT2(!TestNo, 2, 2, 6)
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If

                        If varAttributes(iat).TagString = "XEXTENT" Then
                            If IsNull(!EXTENT) = True Then
                                varAttributes(iat).textString = ""
                            Else
                                varAttributes(iat).textString = !EXTENT
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If
                        If varAttributes(iat).TagString = "XFIGURE" Then
                            If IsNull(!FIGURE) = True Then
                                varAttributes(iat).textString = ""
                            Else
                                varAttributes(iat).textString = !FIGURE
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If
                        If varAttributes(iat).TagString = "XNOTE" Then
                            If IsNull(!NOTE) = True Then
                                varAttributes(iat).textString = ""
                            Else
                            '********** pre 2014 **********
                            '    If Len(!NOTE) > 20 Then
                            '        varAttributes(iat).TextString = Left(!NOTE, 19)
                                
                            '        Note2 = Mid(!NOTE, 20, Len(!NOTE))
                            '    Else
                            '        varAttributes(iat).TextString = !NOTE
                            '    Debug.Print varAttributes(iat).TextString
                            '    End If
                            '********** pre 2014 **********
                                varAttributes(iat).textString = SEPTEXT2(!NOTE, 1, 2, 22)
                            End If
                        End If
                        If varAttributes(iat).TagString = "XNOTE2" Then
                            If IsNull(!NOTE) = True Then
                                varAttributes(iat).textString = ""
                            Else
                                'varAttributes(iat).TextString = Note2
                                varAttributes(iat).textString = SEPTEXT2(!NOTE, 2, 2, 22)
                                Debug.Print varAttributes(iat).textString
                            End If
                        End If
                        
                    Next iat
                    mKey = !Key
        TblPT(1) = TblPT(1) - 20
PARTINEXT:
        .MoveNext
    Next i
End With
'Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                             (tblPT, "XHDWRFOOT02", 1, 1, 1, 0)
RstParti.Close
Set DB = Nothing
ExitPartition:
frmDoor.Hide

End Sub

Private Sub WindSchType_Click()

End Sub

Private Sub writeWindow_Click()
seleproj_Change
End Sub
