Attribute VB_Name = "WriteNotes"
Function WriteCSILabel(TblPT As Variant, mCSI As String, mWKCSINAME As String)
Dim CsiPos(0 To 2) As Variant
Dim CSILabel As String
Dim width As Integer
Dim textObj As AcadText
Dim mtextObj As AcadMText

CsiPos(0) = TblPT(0) - 120
CsiPos(1) = TblPT(1)
CsiPos(2) = TblPT(2)
width = 48 '4' X 12
CSILabel = mCSI & "-" & mWKCSINAME
Set mtextObj = ThisDrawing.ModelSpace.AddMText(CsiPos, width, CSILabel)
    mtextObj.StyleName = "R"
    mtextObj.height = 4
    mtextObj.Update

End Function
Function WRITEWORKBY(mWkID As Variant)

    
Dim ctl As Control
Dim varItm As Variant
Dim dbs As Database, rst As Object
Dim strSQL, tmpWorkby As String
    ' Return reference to current database.
Set dbs = OpenDatabase("\\jwja-svr-10\Jdb\db_est\X_EST_2013_001.mdb")
    ' Open recordset on Employees table.
    'Set rst = dbs.OpenRecordset("ASSit", dbOpenDynaset)
Debug.Print mWkID
strSQL = "SELECT connwkby.connwkby_id, connwkby.wkby, connwkby.wkid, " _
& "connwkby.projwkbyno, connwkby.projwkbynote, workby.workbytype " _
& "FROM workby INNER JOIN connwkby ON workby.workbyid = connwkby.wkby " _
& "WHERE (((connwkby.wkid)=" & mWkID & "))" _
& "ORDER BY connwkby.projwkbyno;"
'strSQL = "SELECT workby.wkid, workby.workbyid, workby.workbyno, workby.workbytype, workby.workbynote" _
& "FROM workby " _
& "ORDER BY workby.workbyno;"
Set rst = dbs.OpenRecordset(strSQL)
mReccount = rst.RecordCount
If mReccount = 3 Then
    Debug.Print "Got It!!!!!!!!!!!!!!!!!!!!!!!!!!"
End If
Debug.Print "FIRST RECCOUNT: " & mReccount
If mReccount = "" Then
    Debug.Print "********************Exiting for mrecount = ""!"""""""""""
    GoTo doaction:
End If
If IsNull(mReccount) = True Then
    Debug.Print "********************Exiting for mrecount = null!"""""""""""
    'printAction
    'Exit Sub
    GoTo doaction
End If
If mReccount = 1 Then
    With rst
        WRITEWORKBY = !workbytype
    End With
    GoTo doaction
End If
With rst
    ' Populate Recordset and print number of records.
    '.MoveLast
    '.MoveFirst
    Debug.Print "  Number of records - Workby = " & _
    .RecordCount
    'intRec = .RecordCount
    intRec = mReccount
    If intRec < 1 Then
        'Exit Sub
        GoTo doaction:
    End If
    If intRec = 1 Then
        tmpWorkby = !workbytype
        'Exit Sub
        Debug.Print "********************Exiting for mrecount = 1!"""""""""""
        WRITEWORKBY = tmpWorkby
        GoTo doaction:
    End If
    '.MoveFirst
    If intRec > 1 Then
        tmpWorkby = ""
        For x = 0 To intRec - 1
            tmpWorkby = tmpWorkby & !workbytype & "/"
            Debug.Print "Count i: and tmpWorkby: " & x & ": " & tmpWorkby & "<<<<<<<<<<-------"
            .MoveNext
        Next x
        nLen = Len(tmpWorkby)
        tmpWorkby = Mid(tmpWorkby, 1, nLen - 1)
        'Me.mWorkby = tmpWorkby
        WRITEWORKBY = tmpWorkby
        
    Else
        tmpWorkby = !workbytype
        nLen = Len(tmpWorkby)
        tmpWorkby = Mid(tmpWorkby, 1, nLen - 1)
        Debug.Print tmpWorkby
        'Me.mWorkby = tmpWorkby
         WRITEWORKBY = tmpWorkby
    End If
End With
doaction:
'printAction (mWkID)
rst.Close
Set dbs = Nothing

End Function
Function printAction(mWkID As Integer)


Dim ctl As Control
Dim varItm As Variant
Dim dbs As Database, rst As Object
Dim strSQL, tmpAction As String
    ' Return reference to current database.
Set dbs = OpenDatabase("\\jwja-svr-10\Jdb\db_est\X_EST_2013_001.mdb")
'DB_EST\X_EST_2013_001.mdb
    ' Open recordset on Employees table.
    'Set rst = dbs.OpenRecordset("ASSit", dbOpenDynaset)
Debug.Print mWkID
'strSQL = "SELECT connwkby.connwkby_id, connwkby.wkby, connwkby.wkid, connwkby.projwkbyno, connwkby.projwkbynote, workby.workbytype " _
& "FROM workby INNER JOIN connwkby ON workby.workbyid = connwkby.wkby " _
& "WHERE (((connwkby.wkid)=" & Me.WKID & "));"
strSQL = "SELECT connwkaction.connwkactionid, connwkaction.actionid, connwkaction.wkid, connwkaction.projwkactionno, workaction.wkaction " _
& "FROM workaction INNER JOIN connwkaction ON workaction.wkactionid = connwkaction.actionid " _
& "WHERE (((connwkaction.WKID) =" & mWkID & ")) " _
& "ORDER BY connwkaction.projwkactionno;"

'strSQL = "SELECT workby.wkid, workby.workbyid, workby.workbyno, workby.workbytype, workby.workbynote" _
& "FROM workby " _
& "ORDER BY workby.workbyno;"
Set rst = dbs.OpenRecordset(strSQL)
With rst
    ' Populate Recordset and print number of records.
    '.MoveLast
    '.MoveFirst
    Debug.Print "  Number of records = Work " & _
    .RecordCount
    intRec = .RecordCount
    If intRec < 1 Then
        rst.Close
        Set dbs = Nothing
        Debug.Print "********************Exiting for intRec < 1!!!!!!!!!!!!!!"
        Exit Function
    End If
    
    '.MoveFirst
    If intRec > 1 Then
        tmpAction = ""
        For x = 0 To intRec - 1
            tmpAction = tmpAction & !wkaction & "/"
            .MoveNext
        Next x
        mLen = Len(tmpAction)
        tmpAction = Mid(tmpAction, 1, mLen - 1)
    Else
        tmpAction = !wkaction
        'Me.mAction = tmpAction 'JIN TEST
        'rst.Close 'JIN TEST
        'Set dbs = Nothing 'JIN TEST
        Debug.Print "********************Exiting for intRec = 1:" & tmpAction & ">>>>>>>"
        Debug.Print tmpAction
    End If
End With
'Me.mAction = tmpAction
printAction = tmpAction
rst.Close
Set dbs = Nothing


End Function
