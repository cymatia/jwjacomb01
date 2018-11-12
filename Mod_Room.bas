Attribute VB_Name = "Mod_Room"
Public Sub writeRoom()
Dim DB As Database
Dim rstRoom As Recordset
Dim strRoom As String
Dim nRm As Integer    'ProjSet no.
Dim iattr_Room As Integer
'Dim rcArray As Variant
Dim TblPT(0 To 2) As Double
Dim tblPT2(0 To 2) As Double
Dim TblEndPT(0 To 2) As Double

Dim txtHt As Double
Dim textObj As AcadText
Dim attributeObj As AcadAttribute
Dim blockRefObj As AcadBlockReference
'Dim BlkObj As Object
'Dim acadPol As AcadLWPolyline 'MOD 9/9/14
'Dim textObj As AcadText 'MOD 9/9/14
Dim txtSecHt As Double
TblPT(0) = 0: TblPT(1) = 0: TblPT(2) = 0
tblPT2(0) = -300: tblPT2(1) = 0: tblPT2(2) = 0
txtHt = 4
txtRow = ""
mode = acAttributeModeVerify

'2'-8" X 1'-4"
nRmNoDWG = "\\jwja-svr-10\drawfile\jvba\X_rmno01.dwg"
nJumpHead = 0
nJumpRmNoROW = 16 '1'-4" X 12
nJumpUnit = 24 '1'-4" X 12
nJumpFl = 48
nJumpSector = 48
Set DB = OpenDatabase("\\jwja-svr-10\jdb\db_misc\door02_01.mdb")
'xProj_no = frmCode.seleproj
xProj_no = frmDoor.seleproj
strSector = "SELECT SECTOR.SECTORID, SECTOR.SECTORORDER, SECTOR.SECTOR, SECTOR.SECTORNOTE, SECTOR.PROJ_NO " _
& "FROM SECTOR " _
& "WHERE (((SECTOR.PROJ_NO)='" & xProj_no & "'))" _
& "order by sector;"
Set rstSector = DB.OpenRecordset(strSector)
rstSector.MoveLast
rstSector.MoveFirst
If rstSector.RecordCount = 0 Then
    GoTo Exit_writeroom
End If
With rstSector 'with 0
    Do While Not rstSector.EOF
    mSectorID = !Sectorid
    
    '******************WRITE SECTOR************************
    '*****************************************************
    'Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                        (tblPT, nRmNoDWG, 1, 1, 1, 0)
    Debug.Print sector
    'Debug.Print !sector.sector
    Debug.Print !sector
    If IsNull(!sector) = False Then
    txtHt = 6
    'Set MTextObj = ThisDrawing.ModelSpace.AddMText(corner, WIDTH, text)
    TblPT(0) = TblPT(0) - 48
    TblPT(1) = TblPT(1) + 12
    Set textObj = ThisDrawing.ModelSpace.AddText _
                ("Sector: " & !sector, TblPT, txtHt)
    TblPT(1) = TblPT(1) - 6
    TblEndPT(0) = TblPT(0) + 96: TblEndPT(1) = TblPT(1): TblEndPT(2) = TblPT(2)
    'Set acadPol = ThisDrawing.ModelSpace.AddLightWeightPolyline(tblPT, tblEndPT)
    ThisDrawing.ModelSpace.AddLine TblPT, TblEndPT
    TblPT(0) = TblPT(0) + 48
    TblPT(1) = TblPT(1) - 6
    End If
    '*****************************************************
        strFl = "SELECT Floor.flID, Floor.floor, Floor.proj_no, Floor.floornote, floor.sectorid " _
        & "FROM floor where floor.sectorid =" & mSectorID & " " _
        & "order by floor;"
        Set rstFl = DB.OpenRecordset(strFl)
        'rstFl.MoveLast
        'rstFl.MoveFirst
        ReccountFl = rstFl.RecordCount
    '******************Jump Sector************************
    '*****************************************************
        If rstFl.RecordCount = 0 Then
            GoTo GotoNextSector
        End If
            With rstFl 'With 1
                rstFl.MoveLast
                rstFl.MoveFirst
                countFl = 0
                Do While Not rstFl.EOF
                ReccountFl = rstFl.RecordCount
                'countFl = countFl + 1
                Debug.Print "Floor: " & countFl & "-" & !floor
                
                mFlID = !flid
                    strUnit = "SELECT Unit.flid, Unit.unitid, Unit.unitorder, Unit.unitno, Unit.unitname, Unit.unitnote " _
                                & "FROM Unit " _
                                & "WHERE (((Unit.flid) =" & mFlID & ")) " _
                                & "ORDER BY Unit.unitorder;"
                    Set rstUnit = DB.OpenRecordset(strUnit)
                    RecCountUnit = rstUnit.RecordCount
        '******************Jump Floor ************************
        '*****************************************************
                    If rstUnit.RecordCount = 0 Then
                        GoTo GotoNextFloor
                    End If
                    rstUnit.MoveLast
                    rstUnit.MoveFirst
                    With rstUnit 'With 2
                    
                    rstUnit.MoveLast
                    rstUnit.MoveFirst
                    countUnit = 0
                    Do While Not rstUnit.EOF
                    munitID = !unitid
                    countUnit = countUnit + 1
                    Debug.Print "Unit: " & countUnit & "-" & !unitno
                    strRm = "SELECT ROOM.RMID, ROOM.RMORDER, ROOM.RMNO, ROOM.RMNAME, ROOM.PDRID, ROOM.NOTE, ROOM.PROJ_NO, " _
                            & "ROOM.FLID, ROOM.UnitID, ROOM.FRMNO " _
                            & "FROM ROOM " _
                            & "WHERE (((room.unitid) =" & munitID & ")) " _
                            & "ORDER BY ROOM.RMORDER;"
                            Set rstRoom = DB.OpenRecordset(strRm)
                            'rstRoom.MoveLast
                            'rstRoom.MoveFirst
        '******************Jump Unit  ************************
        '*****************************************************
                            If rstRoom.RecordCount = 0 Then
                                GoTo GotoNextUnit
                            End If
                            RecCountRm = rstRoom.RecordCount
                            nRmNo = rstRoom.RecordCount
                            With rstRoom 'With 3
                                'For i3 = 0 To nRmNo - 1
                                countRm = 0
                               Do While Not rstRoom.EOF
                                    countRm = countRm + 1
                                    Debug.Print "Room: " & countRm & "-" & !rmno
                                    Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
                                                                 (TblPT, nRmNoDWG, 1, 1, 1, 0)
                                    Dim varAttributes1 As Variant
                                    varAttributes1 = blockRefObj.GetAttributes
                                    For iattr_RMrow = LBound(varAttributes1) To UBound(varAttributes1)
                                        '************ Xrmno ***************************************************************
                                        If varAttributes1(iattr_RMrow).TagString = "XRMNO" Then
                                            If IsNull(!FRMNO) = True Then '++++++++++++++++++++++++++++++++++++++++++
                                                varAttributes1(iattr_RMrow).textString = ""
                                            Else
                                                varAttributes1(iattr_RMrow).textString = !FRMNO
                                                
                                            End If
                                        End If
                                        '************ Xrmname ***************************************************************
                                        If varAttributes1(iattr_RMrow).TagString = "XRMNAME" Then
                                            If IsNull(!FRMNO) = True Then '++++++++++++++++++++++++++++++++++++++++++
                                                varAttributes1(iattr_RMrow).textString = ""
                                            Else
                                                If IsNull(!RMNAME) = True Then
                                                varAttributes1(iattr_RMrow).textString = "CHECK!"
                                                Else
                                                varAttributes1(iattr_RMrow).textString = !RMNAME
                                                End If
                                                
                                            End If
                                        End If
                                    
                                    Next iattr_RMrow
                                    TblPT(1) = TblPT(1) - nJumpRmNoROW
                                    .MoveNext
                                    'Next i3
                                Loop
                                
                            End With 'With 3
                            rstRoom.Close
GotoNextUnit:
                            .MoveNext
                            TblPT(1) = TblPT(1) - nJumpUnit
                       Loop
                
                             End With 'With 2
                                     'rstunit.Close
GotoNextFloor:
                .MoveNext
                TblPT(1) = TblPT(1) - nJumpFl
                Loop
            
            End With 'With 1
GotoNextSector:
    .MoveNext
    TblPT(1) = TblPT(1) - nJumpSector
    
    Loop
End With 'with 0
'rstfl.Close

Set DB = Nothing

Exit_writeroom:
rstSector.Close
'rstRoom.Close
rstUnit.Close
rstFl.Close

'frmCode.HIDE
frmDoor.Hide

Set DB = Nothing
End Sub

