Attribute VB_Name = "SepText"
Function SEPTEXT(xText As String, ATno As Integer, lNo As Integer, tNo As Integer)
Dim strArray() As String
'Dim strData(0) As String
Dim strType() As String
Dim strLen() As Integer
Dim mType, tn As Variant
Dim strTimeStamp, srchChar1, srchChar2, srchChar3, xstrArray As String
Dim pos, txLen, xtNo, xSp, sp, sl, m, n, k, i, j, xstrLen As Integer
Dim blCutLast, blSl, blSp As Boolean

srchChar1 = Chr(32) '32 is space 45 => -, 47 => /
srchChar2 = Chr(45)
srchChar3 = Chr(47)

'strData(0) = xTEXT
pos = 1
sp = 0
sl = 0
If IsNull(xText) = True Then
    Exit Function
End If
Debug.Print Len(xText)
txLen = Len(xText)
If txLen = 0 Then
    Exit Function
End If


If txLen < tNo Then
    If ATno = 1 Then
    'SEPTEXT = ""
        SEPTEXT = xText
        Debug.Print "SEPTEXT @ SMALL tNo:" & SEPTEXT
    Else
        SEPTEXT = ""
    End If
    Exit Function
End If
' *************** Find space
If InStr(xText, srchChar1) <> 0 Then
    blSp = True
    For m = 1 To txLen
        If Mid(xText, m, 1) = srchChar1 Then
            sp = sp + 1
            blSp = True
        End If
    Next m
    ReDim strArray(sp + 1)
    Debug.Print "strArray(sp+1): " & strArray(sp + 1)
    ReDim strLen(sp + 1)
    Debug.Print "strLen(sp+1): " & strLen(sp + 1)
End If


If InStr(xText, srchChar2) <> 0 Then
    blSl = True
    For n = 1 To txLen
        If Mid(xText, n, 1) = srchChar2 Then
            sl = sl + 1
            blSl = True
        End If
    Next n
End If



Debug.Print "Total sp: " & sp & "Total sl: " & sl

xtNo = txLen / lNo 'MAX LINE NO?
xtNo = CInt(xtNo)
Debug.Print "xtNo: " & xtNo
If xtNo > tNo Then 'tNo = max text length variable
    blCutLast = True
End If


If blSp = True Then
    For i = 1 To sp + 1
    If i = lNo Then
        strArray(i) = Mid(xText, 1, txLen)
        xText = ""
        GoTo checkcomb
    Else
        xSp = InStr(xText, srchChar1)
        If xSp <> 0 Then
           strArray(i) = Mid(xText, 1, xSp)
        Else
           strArray(i) = Mid(xText, 1, txLen)
        End If
        
        strLen(i) = Len(strArray(i))
        xText = Mid(xText, xSp + 1, txLen)
    End If
    Debug.Print i & ": " & strArray(i)
    Next i
End If
checkcomb:
Debug.Print i & "<<<< i"
'If i - 1 > lNo Then ' combine text
If i > lNo Then  ' combine text
    For j = 1 To sp + 1
        If j = 1 Then
            xstrArray = ""
            xstrLen = 0
            GoTo NEXTj
        End If
        If xstrLen + strLen(i) > tNo Then
            strArray(i) = xstrArray & strArray(i)
        End If
NEXTj:
        
    Next j
    xstrArray = strArray(i)
        xstrLen = strLen(i)
End If
setText:
For k = 1 To lNo
    If k = ATno Then
        If k = lNo And blCutLast = True Then
            SEPTEXT = Mid(strArray(k), 1, tNo)
            Exit Function
        End If
    
    SEPTEXT = strArray(k)
    End If
Next k
Erase strArray, strLen, strType 'Jin added this line 10/21/12
Debug.Print SEPTEXT

End Function
Function SEPTEXT2(xText As String, ATno As Integer, lNo As Integer, tNo As Integer)
Dim strArray() As String
'Dim strData(0) As String
Dim strType() As String
Dim strLen() As Integer
Dim posSpace() As Integer
Dim posSlash() As Integer
Dim mType, tn As Variant
Dim strTimeStamp, srchChar1, srchChar2, srchChar3, xstrArray As String
Dim pos, txLen, xtNo, xSp, sp, sl, m, n, k, i, j, xstrLen, position As Integer
Dim blCutLast, blSl, blSp As Boolean

srchChar1 = Chr(32) '32 is space 45 => -, 47 => /
srchChar2 = Chr(47)
srchChar3 = Chr(45)
'++++++++++++++++++++++++++++++++
'xTEXT = Forms!frmtest!Text
'ATno = Forms!frmtest!lineno
'lNo = Forms!frmtest!linetotal
'tNo = Forms!frmtest!chrperline
'++++++++++++++++++++++++++++++++


'strData(0) = xTEXT
pos = 1
sp = 0
sl = 0
If IsNull(xText) = True Then
    Exit Function
End If
If xText = "" Then
    Exit Function
End If

Debug.Print Len(xText)
txLen = Len(xText)
If txLen = 0 Then
    Exit Function
End If


If txLen < tNo Then
    If ATno = 1 Then
    'SEPTEXT = ""
        SEPTEXT2 = xText
        Debug.Print "SEPTEXT @ SMALL tNo:" & SEPTEXT2
    Else
        SEPTEXT2 = ""
    End If
    Exit Function
End If
'****************************************************************
'***************  Find space      *******************************
'****************************************************************
If InStr(xText, srchChar1) <> 0 Then
    blSp = True
    For m = 1 To txLen
        If Mid(xText, m, 1) = srchChar1 Then
            sp = sp + 1
           ' Debug.Print strArray(m)
            'strArray() = strArray(sp, m)
            blSp = True
        End If
    Next m
    Debug.Print blSp
 
    
    xtNo = txLen / lNo 'MAX LINE NO?
    xtNo = CInt(xtNo)
    Debug.Print "xtNo: " & xtNo
    If xtNo > tNo Then 'tNo = max text length variable
        blCutLast = True
    End If
    ReDim strArray(lNo + 1)
    Debug.Print "strArray(lNo): " & strArray(lNo)
    ReDim strLen(sp + 1)
    Debug.Print "strLen(sp+1): " & strLen(sp + 1)
    ReDim posSpace(sp + 1)
End If
'****************************************************************
'***************  Find slash      *******************************
'****************************************************************
If InStr(xText, srchChar1) = 0 And InStr(xText, srchChar2) <> 0 Then
    blSl = True
    For n = 1 To txLen
        If Mid(xText, n, 1) = srchChar2 Then
            sl = sl + 1
            blSl = True
        End If
    Next n
    ReDim strArray(lNo + 1)
    Debug.Print "strArray(lNo): " & strArray(lNo)
    ReDim strLen(sl + 1)
    Debug.Print "strLen(sl+1): " & strLen(sl + 1)
    ReDim posSlash(sl + 1)
    Debug.Print posSlash(0)
End If
'****************************************************************
'***************  End Find  *******************************
'****************************************************************


Debug.Print "Total sp: " & sp & " Total sl: " & sl
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
If blSp = True Then
    'For i = 1 To sp + 1
    For i = 1 To lNo
        If i = lNo Then
            strArray(i) = Mid(xText, 1, txLen)
            Debug.Print strArray(i)
            xText = ""
            GoTo checkcomb
        Else
            xSp = InStr(xText, srchChar1)
            
            Dim cntSpace As Integer
                For m = 1 To txLen
                    
                        If Mid(xText, m, 1) = srchChar1 Then
                        
                            cntSpace = cntSpace + 1
                            position = m
                            posSpace(cntSpace) = m
                            Debug.Print posSpace(cntSpace)
                            'strArray() = strArray(sp, m)
                            blSp = True
                        End If
                    If m >= tNo Then
                    Exit For
                    End If
                Next m
            If posSpace(cntSpace) = 0 Then
            strArray(i) = Mid(xText, 1, tNo)
            Debug.Print strArray(i)
            xText = Mid(xText, tNo + 1, txLen)
            GoTo skipfor
            End If
            If xSp <> 0 Then 'Whether space exist, if it does then
               'strArray(i) = Mid(xTEXT, 1, xSp)
               strArray(i) = Mid(xText, 1, posSpace(cntSpace))
               Debug.Print strArray(i)
            Else
               strArray(i) = Mid(xText, 1, txLen)
            End If
            
            'strLen(i) = Len(strArray(i)) '+++++++++++++++++++++++++++++Subscription Error
            'xTEXT = Mid(xTEXT, xSp + 1, txLen)
            xText = Mid(xText, posSpace(cntSpace) + 1, txLen)
        End If
skipfor:
    Debug.Print i & ": " & strArray(i)
    Next i
    Debug.Print i & ": " & strArray(i)
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

If blSp = False And blSl = True Then
    For i = 1 To lNo
        If i = lNo Then
            strArray(i) = Mid(xText, 1, txLen)
            Debug.Print strArray(i)
            xText = ""
            GoTo checkcomb
        Else
            xSl = InStr(xText, srchChar1)
            
            Dim cntSlash As Integer
            For m = 1 To txLen
                
                    If Mid(xText, m, 1) = srchChar2 Then
                    
                        cntSlash = cntSlash + 1
                        position = m
                        posSlash(cntSlash) = m
                        Debug.Print posSlash(cntSlash)
                        'strArray() = strArray(sp, m)
                        blSl = True
                    End If
                If m >= tNo Then
                    Exit For
                End If
            Next m
            If posSlash(cntSlash) = 0 Then
                strArray(i) = Mid(xText, 1, tNo)
                Debug.Print strArray(i)
                xText = Mid(xText, tNo + 1, txLen)
                GoTo skipfor2
            End If
            If xSl <> 0 Then 'Whether space exist, if it does then
               'strArray(i) = Mid(xTEXT, 1, xSp)
               strArray(i) = Mid(xText, 1, posSlash(cntSlash))
               Debug.Print strArray(i)
            Else
               strArray(i) = Mid(xText, 1, txLen)
            End If
            
            strLen(i) = Len(strArray(i))
            'xTEXT = Mid(xTEXT, xSp + 1, txLen)
            xText = Mid(xText, posSlash(cntSlash) + 1, txLen)
            Debug.Print xText
        End If
skipfor2:
    Debug.Print i & ": " & strArray(i)
    Next i
    Debug.Print i & ": " & strArray(i)
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


'Debug.Print i & ": " & strArray(i)
'****************************CheckComb ***********************
checkcomb:
Debug.Print i & "<<<< i line numbers"
 
'If i - 1 > lNo Then ' combine text
If i > lNo Then  ' combine text
ReDim strArray(k + 1) As String


    For j = 1 To sp + 1
        If j = 1 Then
            xstrArray = ""
            xstrLen = 0
            GoTo NEXTj
        End If
        If xstrLen + strLen(i) > tNo Then
            strArray(i) = xstrArray & strArray(i)
        End If
NEXTj:
        
    Next j
    xstrArray = strArray(i)
        xstrLen = strLen(i)
End If
'**************************** Set Text ***********************

setText:
For k = 1 To lNo
    If k = ATno Then
        'If k = lNo And blCutLast = True Then
        If k = lNo Then
            'SEPTEXT2 = Mid(strArray(k), 1, tNo)
            SEPTEXT2 = strArray(k)
            Exit Function
        End If
    'Debug.Print "strArray(k):" & k & " : " & strArray(k)
    SEPTEXT2 = strArray(k)
    'SEPTEXT2 = "Error"
        If Err.Description = "Subscription error" Then
        MsgBox strArray(i) & " could not be set."
        Exit Function
    End If

    'SEPTEXT2 = "first of range"
    End If
Next k
Erase strArray, strLen, strType 'Jin added this line 10/21/12
Debug.Print SEPTEXT2



End Function

Function SEPTEXT3(xText As String, ATno As Integer, lNo As Integer, tNo As Integer)
Dim strArray() As String
'Dim strData(0) As String
Dim strType() As String
Dim strLen() As Integer
Dim mType, tn As Variant
Dim strTimeStamp, srchChar1, srchChar2, srchChar3, xstrArray As String
Dim pos, txLen, xtNo, xSp, sp, sl, m, n, k, i, j, xstrLen As Integer
Dim blCutLast, blSl, blSp As Boolean

srchChar1 = Chr(32) '32 is space 45 => -, 47 => /
srchChar2 = Chr(45)
srchChar3 = Chr(47)

'strData(0) = xTEXT
pos = 1
sp = 0
sl = 0
If IsNull(xText) = True Then
Exit Function
End If
Debug.Print Len(xText)
txLen = Len(xText)
If txLen = 0 Then
Exit Function
End If


If txLen < tNo Then
    If ATno = 1 Then
    SEPTEXT3 = xText
    Else
    SEPTEXT3 = ""
    End If
    Exit Function
End If

If InStr(xText, srchChar1) <> 0 Then
    blSp = True
    For m = 1 To txLen
        If Mid(xText, m, 1) = srchChar1 Then
            sp = sp + 1
            blSp = True
        End If
    Next m
    ReDim strArray(sp + 1)
    ReDim strLen(sp + 1)
End If


If InStr(xText, srchChar2) <> 0 Then
    blSl = True
    For n = 1 To txLen
        If Mid(xText, n, 1) = srchChar2 Then
            sl = sl + 1
            blSl = True
        End If
    Next n
End If



Debug.Print "Total sp: " & sp & "Total sl: " & sl

xtNo = txLen / lNo 'MAX LINE NO?
xtNo = CInt(xtNo)

If xtNo > tNo Then
    blCutLast = True
End If


If blSp = True Then
    For i = 1 To sp + 1
    If i = lNo Then
    strArray(i) = Mid(xText, 1, txLen)
    xText = ""
    GoTo checkcomb
    Else
        xSp = InStr(xText, srchChar1)
        If xSp <> 0 Then
           strArray(i) = Mid(xText, 1, xSp)
        Else
           strArray(i) = Mid(xText, 1, txLen)
        End If
        
        strLen(i) = Len(strArray(i))
        xText = Mid(xText, xSp + 1, txLen)
    End If
    Debug.Print i & ": " & strArray(i)
    Next i
End If
checkcomb:
Debug.Print i & "<<<< i"
'If i - 1 > lNo Then ' combine text
If i > lNo Then  ' combine text
    For j = 1 To sp + 1
        If j = 1 Then
            xstrArray = ""
            xstrLen = 0
            GoTo NEXTj
        End If
        If xstrLen + strLen(i) > tNo Then
            strArray(i) = xstrArray & strArray(i)
        End If
NEXTj:
        
    Next j
    xstrArray = strArray(i)
        xstrLen = strLen(i)
End If
setText:
For k = 1 To lNo
    If k = ATno Then
        If k = lNo And blCutLast = True Then
            SEPTEXT3 = Mid(strArray(k), 1, tNo)
            Exit Function
        End If
    
    SEPTEXT3 = strArray(k)
    End If
Next k

Debug.Print SEPTEXT3

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   SEPTEXT2 FROM MOD_SCHED 9/8/14
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function SEPTEXT22(xText As String, ATno As Integer, lNo As Integer, tNo As Integer)
Dim strArray() As String
'Dim strData(0) As String
Dim strType() As String
Dim strLen() As Integer
Dim posSpace() As Integer
Dim posDash() As Integer
Dim mType, tn As Variant
Dim strTimeStamp, srchChar1, srchChar2, srchChar3, xstrArray As String
Dim pos, txLen, xtNo, xSp, xDs, sp, sl, ds, m, n, k, i, j, xstrLen, position As Integer
Dim blCutLast, blSl, blSp, blDs As Boolean

srchChar1 = Chr(32) '32 is space 45 => -, 47 => /
srchChar2 = Chr(47)
srchChar3 = Chr(45)

'strData(0) = xTEXT
pos = 1
sp = 0
sl = 0
If IsNull(xText) = True Then
    Exit Function
End If
If xText = "" Then
    Exit Function
End If

Debug.Print Len(xText)
txLen = Len(xText)
If txLen = 0 Then
    Exit Function
End If


If txLen < tNo Then
    If ATno = 1 Then
    'SEPTEXT = ""
        SEPTEXT22 = xText
        Debug.Print "SEPTEXT @ SMALL tNo:" & SEPTEXT22
    Else
        SEPTEXT22 = ""
    End If
    Exit Function
End If
' *************** Find space ***************
If InStr(xText, srchChar1) <> 0 Then
    blSp = True
    For m = 1 To txLen
        If Mid(xText, m, 1) = srchChar1 Then
            sp = sp + 1
           ' Debug.Print strArray(m)
            'strArray() = strArray(sp, m)
            blSp = True
        End If
    Next m
    
    
xtNo = txLen / lNo 'MAX LINE NO?
xtNo = CInt(xtNo)
Debug.Print "xtNo: " & xtNo
If xtNo > tNo Then 'tNo = max text length variable
    blCutLast = True
End If
    
    
    
    ReDim strArray(lNo + 1)
    Debug.Print "strArray(lNo): " & strArray(lNo)
    ReDim strLen(sp + 1)
    Debug.Print "strLen(sp+1): " & strLen(sp + 1)
    ReDim posSpace(sp + 1)
End If

' *************** Find slash ***************

If InStr(xText, srchChar2) <> 0 Then
    blSl = True
    For n = 1 To txLen
        If Mid(xText, n, 1) = srchChar2 Then
            sl = sl + 1
            blSl = True
        End If
    Next n
End If



Debug.Print "Total sp: " & sp & " Total sl: " & sl
' *************** Find DASH ***************
If InStr(xText, srchChar3) <> 0 Then
    blDs = True
    For n = 1 To txLen
        If Mid(xText, n, 1) = srchChar3 Then
            ds = ds + 1
            blDs = True
        End If
    Next n
End If
'**************************************
'**********************************************************************
'**********************************************************************
'************************* blSp ****************************
'**********************************************************************
'**********************************************************************
If blSp = True Then
    For i = 1 To sp + 1
        If i = lNo Then
            strArray(i) = Mid(xText, 1, txLen)
            Debug.Print strArray(i)
            xText = ""
            GoTo checkcomb
        Else
            xSp = InStr(xText, srchChar1)
            Dim cntSpace As Integer
                For m = 1 To txLen
                    
                        If Mid(xText, m, 1) = srchChar1 Then
                        
                            cntSpace = cntSpace + 1
                            position = m
                            posSpace(cntSpace) = m
                            Debug.Print posSpace(cntSpace)
                            'strArray() = strArray(sp, m)
                            blSp = True
                        End If
                    If m >= tNo Then
                    Exit For
                    End If
                Next m
            
            If xSp <> 0 Then
               'strArray(i) = Mid(xTEXT, 1, xSp)
               strArray(i) = Mid(xText, 1, posSpace(cntSpace))
               Debug.Print strArray(i)
            Else
               strArray(i) = Mid(xText, 1, txLen)
            End If
            
            strLen(i) = Len(strArray(i))
            'xTEXT = Mid(xTEXT, xSp + 1, txLen)
            xText = Mid(xText, posSpace(cntSpace) + 1, txLen)
        End If
    Debug.Print i & ": " & strArray(i)
    Next i
End If

'**********************************************************************
'**********************************************************************
'************************* blDs ****************************
'**********************************************************************
'**********************************************************************


If blDs = True Then
    For i = 1 To ds + 1
        If i = lNo Then
            strArray(i) = Mid(xText, 1, txLen)
            Debug.Print strArray(i)
            xText = ""
            GoTo checkcomb
        Else
            xDs = InStr(xText, srchChar3)
            Dim cntDash As Integer
            ReDim posDash(ds)
            
                For m = 1 To txLen
                    
                        If Mid(xText, m, 1) = srchChar3 Then 'If space or dash
                        
                            cntDash = cntDash + 1
                            position = m
                            posDash(cntDash) = m
                            Debug.Print posDash(cntDash)
                            'strArray() = strArray(sp, m)
                            blDs = True
                        End If
                    If m >= tNo Then
                    Exit For
                    End If
                Next m
            
            If xDs <> 0 Then
               'strArray(i) = Mid(xTEXT, 1, xSp)
               strArray(i) = Mid(xText, 1, posDash(cntDash))
               Debug.Print strArray(i)
            Else
               strArray(i) = Mid(xText, 1, txLen)
            End If
            
            strLen(i) = Len(strArray(i))
            'xTEXT = Mid(xTEXT, xSp + 1, txLen)
            xText = Mid(xText, posSpace(cntDash) + 1, txLen)
        End If
    Debug.Print i & ": " & strArray(i)
    Next i
End If

'**********************************************************************
'**********************************************************************
'**********************************************************************
'**********************************************************************

'****************************CheckComb ***********************
checkcomb:
Debug.Print i & "<<<< i"
'If i - 1 > lNo Then ' combine text
If i > lNo Then  ' combine text
ReDim strArray(k + 1) As String


    For j = 1 To sp + 1
        If j = 1 Then
            xstrArray = ""
            xstrLen = 0
            GoTo NEXTj
        End If
        If xstrLen + strLen(i) > tNo Then
            strArray(i) = xstrArray & strArray(i)
        End If
NEXTj:
        
    Next j
    xstrArray = strArray(i)
        xstrLen = strLen(i)
End If
'**************************** Set Text ***********************

setText:
'***************** MOD 3/15/13 *****************'
'If strArray(i) = "" Then
'Exit Function
'End If
'***************** MOD 3/15/13 *****************'
'***************** MOD 2/19/14 *****************'
If SEPTEXT22 = "" Then
Exit Function
End If
If IsNull(SEPTEXT22) = True Then
Exit Function
End If
'***************** E MOD 2/19/14 *****************'
For k = 0 To lNo - 1
    If k = ATno Then
        If k = lNo And blCutLast = True Then
            SEPTEXT22 = Mid(strArray(k), 1, tNo)
            Exit Function
        End If
    
    SEPTEXT22 = strArray(k) 'WAS (K)
    'SEPTEXT2 = "second of range"
    End If
Next k
Erase strArray, strLen, strType 'Jin added this line 10/21/12
Debug.Print SEPTEXT22

End Function

Function SEPTEXT01(xText As String, ATno As Integer, lNo As Integer, tNo As Integer) 'MOD_SCHED
Dim strArray() As String
'Dim strData(0) As String
Dim strType() As String
Dim strLen() As Integer
Dim mType, tn As Variant
Dim strTimeStamp, srchChar1, srchChar2, srchChar3, xstrArray As String
Dim pos, txLen, xtNo, xSp, sp, sl, m, n, k, i, j, xstrLen As Integer
Dim blCutLast, blSl, blSp As Boolean

srchChar1 = Chr(32) '32 is space 45 => -, 47 => /
srchChar2 = Chr(45)
srchChar3 = Chr(47)

'strData(0) = xTEXT
pos = 1
sp = 0
sl = 0
If IsNull(xText) = True Then
    Exit Function
End If
Debug.Print Len(xText)
txLen = Len(xText)
If txLen = 0 Then
    Exit Function
End If


If txLen < tNo Then
    If ATno = 1 Then
    'SEPTEXT = ""
        SEPTEXT01 = xText
        Debug.Print "SEPTEXT01 @ SMALL tNo:" & SEPTEXT01
    Else
        SEPTEXT01 = ""
    End If
    Exit Function
End If
' *************** Find space
If InStr(xText, srchChar1) <> 0 Then
    blSp = True
    For m = 1 To txLen
        If Mid(xText, m, 1) = srchChar1 Then
            sp = sp + 1
            blSp = True
        End If
    Next m
    ReDim strArray(sp + 1)
    Debug.Print "strArray(sp+1): " & strArray(sp + 1)
    ReDim strLen(sp + 1)
    Debug.Print "strLen(sp+1): " & strLen(sp + 1)
End If


If InStr(xText, srchChar2) <> 0 Then
    blSl = True
    For n = 1 To txLen
        If Mid(xText, n, 1) = srchChar2 Then
            sl = sl + 1
            blSl = True
        End If
    Next n
End If



Debug.Print "Total sp: " & sp & "Total sl: " & sl

xtNo = txLen / lNo 'MAX LINE NO?
xtNo = CInt(xtNo)
Debug.Print "xtNo: " & xtNo
If xtNo > tNo Then 'tNo = max text length variable
    blCutLast = True
End If


If blSp = True Then
    For i = 1 To sp + 1
    If i = lNo Then
        strArray(i) = Mid(xText, 1, txLen)
        xText = ""
        GoTo checkcomb
    Else
        xSp = InStr(xText, srchChar1)
        If xSp <> 0 Then
           strArray(i) = Mid(xText, 1, xSp)
        Else
           strArray(i) = Mid(xText, 1, txLen)
        End If
        
        strLen(i) = Len(strArray(i))
        xText = Mid(xText, xSp + 1, txLen)
    End If
    Debug.Print i & ": " & strArray(i)
    Next i
End If
checkcomb:
Debug.Print i & "<<<< i"
'If i - 1 > lNo Then ' combine text
If i > lNo Then  ' combine text
    For j = 1 To sp + 1
        If j = 1 Then
            xstrArray = ""
            xstrLen = 0
            GoTo NEXTj
        End If
        If xstrLen + strLen(i) > tNo Then
            strArray(i) = xstrArray & strArray(i)
        End If
NEXTj:
        
    Next j
    xstrArray = strArray(i)
        xstrLen = strLen(i)
End If
setText:
For k = 1 To lNo
    If k = ATno Then
        If k = lNo And blCutLast = True Then
            SEPTEXT01 = Mid(strArray(k), 1, tNo)
            Exit Function
        End If
    
    SEPTEXT01 = strArray(k)
    End If
Next k
Erase strArray, strLen, strType 'Jin added this line 10/21/12
Debug.Print SEPTEXT01

End Function

