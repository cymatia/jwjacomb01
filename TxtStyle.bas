Attribute VB_Name = "TxtStyle"
Sub CreateCNstyle()
Dim newText As AcadTextStyle
Dim blnFoundText As Boolean
Dim blnFoundTitle As Boolean
Dim blnFoundSub As Boolean
   
'reset blnFound
blnFoundText = False
blnFoundTitle = False
blnFoundSub = False

    For i = 0 To ThisDrawing.TextStyles.Count - 1  'make the styles for the form
        Debug.Print ThisDrawing.TextStyles(i).Name
        If ThisDrawing.TextStyles(i).Name = "CN" Then
            blnFoundText = True
            Exit Sub
        End If
        
    Next i

Set newText = ThisDrawing.TextStyles.Add("CN")

typeface = "Courier new"
newText.SetFont typeface, Bold, Italic, Charset, PitchandFamily

End Sub
Sub CreateARstyle()
Dim newText As AcadTextStyle
Dim blnFoundText As Boolean
Dim blnFoundTitle As Boolean
Dim blnFoundSub As Boolean
   
'reset blnFound
blnFoundText = False
blnFoundTitle = False
blnFoundSub = False

    For i = 0 To ThisDrawing.TextStyles.Count - 1  'make the styles for the form
        Debug.Print ThisDrawing.TextStyles(i).Name
        If ThisDrawing.TextStyles(i).Name = "AR" Then
            blnFoundText = True
            Exit Sub
        End If
        
    Next i

Set newText = ThisDrawing.TextStyles.Add("AR")
typeface = "arial"

newText.SetFont typeface, Bold, Italic, Charset, PitchandFamily

End Sub

Sub CreateRstyle()
Dim newText As AcadTextStyle
Dim blnFoundText As Boolean
Dim blnFoundTitle As Boolean
Dim blnFoundSub As Boolean
   
'reset blnFound
blnFoundText = False
blnFoundTitle = False
blnFoundSub = False

    For i = 0 To ThisDrawing.TextStyles.Count - 1  'make the styles for the form
        Debug.Print ThisDrawing.TextStyles(i).Name
        If ThisDrawing.TextStyles(i).Name = "R" Then
            blnFoundText = True
            Exit Sub
        End If
        
    Next i

Set newText = ThisDrawing.TextStyles.Add("R")
typeface = "romans"

newText.SetFont typeface, Bold, Italic, Charset, PitchandFamily

End Sub
Sub CreateRstyle1() 'FROM MOD_SCHED
Dim newText As AcadTextStyle
Dim blnFoundText As Boolean
Dim blnFoundTitle As Boolean
Dim blnFoundSub As Boolean
   
'reset blnFound
blnFoundText = False
blnFoundTitle = False
blnFoundSub = False

    For i = 0 To ThisDrawing.TextStyles.Count - 1  'make the styles for the form
        Debug.Print ThisDrawing.TextStyles(i).Name
        If ThisDrawing.TextStyles(i).Name = "R" Then
            blnFoundText = True
            Exit Sub
        End If
        
    Next i

Set newText = ThisDrawing.TextStyles.Add("R")
typeface = "romans"

newText.SetFont typeface, Bold, Italic, Charset, PitchandFamily

End Sub

Sub CreateARstyle1() 'MOD_SPEC
Dim newText As AcadTextStyle
Dim blnFoundText As Boolean
Dim blnFoundTitle As Boolean
Dim blnFoundSub As Boolean
   
'reset blnFound
blnFoundText = False
blnFoundTitle = False
blnFoundSub = False

    For i = 0 To ThisDrawing.TextStyles.Count - 1  'make the styles for the form
        Debug.Print ThisDrawing.TextStyles(i).Name
        If ThisDrawing.TextStyles(i).Name = "AR" Then
            blnFoundText = True
            Exit Sub
        End If
        
    Next i

Set newText = ThisDrawing.TextStyles.Add("AR")
typeface = "arial"

newText.SetFont typeface, Bold, Italic, Charset, PitchandFamily

End Sub


