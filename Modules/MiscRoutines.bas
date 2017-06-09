Attribute VB_Name = "MiscRoutines"
Function appVer() As Integer

    If UCase(Right(ThisWorkbook.Name, 4)) = ".XLA" Or UCase(Right(ThisWorkbook.Name, 4)) = ".XLS" Then 'running as a 2003 add-in
        appVer = Application.WorksheetFunction.Min(Application.Version, 11)
    Else
        appVer = Application.Version
    End If
End Function
Public Sub renameChartObject()
'sample code to change the name of a chart object, as a bookmark indicator
    ActiveSheet.ChartObjects("zActualsVPlanChart").Name = "ActualsVPlanChart"
End Sub
Public Sub nameEmbeddedObject(Optional control As Object) 'IRibbonControl
Dim myShape As Object
Dim objName As String
Dim bChart As Boolean
Dim chkObj As Long

    If Application.Workbooks.Count = 0 Then
        MsgBox "No files open to process"
        Exit Sub
    End If
    
    If ActiveSheet.Type <> xlWorksheet Then
        MsgBox "You can only rename this chart by renaming the tab." & vbCrLf & _
            "Use the 'Name Embedded Chart' function for charts embedded in worksheets.", vbOKOnly, "Not for use on Chart Tabs..."
    Else
    
        On Error Resume Next
        Set myShape = ActiveChart
        If Not myShape Is Nothing Then
            bChart = True
        Else
            bChart = False
            chkObj = Selection.ShapeRange.Count
            If chkObj > 0 Then
                Set myShape = Selection
            Else
                Set myShape = Nothing
            End If
        End If
        On Error GoTo 0
        
        If myShape Is Nothing Then
            MsgBox "Please select shape/image/chart first"
        Else
            objName = InputBox("Enter new name for selected chart", myShape.Name)
            If objName <> "" And objName <> myShape.Name Then
                On Error Resume Next
                If bChart Then
                    myShape.Parent.Name = objName 'name chart
                Else
                    myShape.Name = objName 'name image/shape
                End If
                If Err.Number <> 0 Then MsgBox "The name: " & objName & " is not valid."
                On Error GoTo 0
            End If
        End If
    End If
End Sub
Public Function getPathFromPathFName(strPath As String) As String
    getPathFromPathFName = Left(strPath, Len(strPath) - InStr(StrReverse(strPath), "\") + 1)
End Function
Public Function getFileExt(fName As String) As String
Dim i As Integer
    
    i = InStr(StrReverse(fName), ".")
    getFileExt = StrReverse(Left(StrReverse(fName), i))
    
End Function
Public Sub unHideAllNames()
Dim myName As Name

    For Each myName In Application.Names
        myName.Visible = True
    Next myName
    
End Sub
Public Function myEvaluate(myName As Variant) As Variant
Dim tmpName As Name
Dim found As Boolean
Dim strSearch As String
Dim chkRange As Range

'Finds an existing Range Name scoped at the Activesheet level, then at the Workbook level, until found, returning a null string if not found
'Once found, returns the evaluation of that name via the Evaluate method, as a Range Object or String.
    
    For Each tmpName In ActiveSheet.Names
        myQuote_char = QUOTE_CHAR
        
        strSearch = UCase(myQuote_char & ActiveSheet.Name & myQuote_char & "!" & myName)
    
        If UCase(tmpName.Name) = strSearch Then
            On Error Resume Next
            Set chkRange = ActiveSheet.Names(myName).RefersToRange
            If Err.Number = 0 Then
                Set myEvaluate = ActiveSheet.Names(myName).RefersToRange
            Else
                myEvaluate = Evaluate(myName)
            End If
            On Error GoTo 0
            found = True
            Exit For
        End If
    Next tmpName
    
    If Not found Then
        For Each tmpName In ActiveWorkbook.Names
        If UCase(tmpName.Name) = UCase(myName) Then
            On Error Resume Next
            Set chkRange = ActiveWorkbook.Names(myName).RefersToRange
            If Err.Number = 0 Then
                Set myEvaluate = ActiveWorkbook.Names(myName).RefersToRange
            Else
                myEvaluate = Evaluate(myName)
            End If
            On Error GoTo 0
            found = True
            Exit For
        End If
        Next tmpName
    End If
    
    If Not found Then myEvaluate = ""
    
End Function
'Source:  Adapted from http://www.codeforexcelandoutlook.com/excel-vba/validate-filenames/
'Interestingly enough, the original code was incorrect.  After referring to Ross McLean's original
'blog, the pattern was corrected, and the Not was added to correctly return results.  Also, in line with disallowed
'Windows filename characters, the | character was added.
'-----------------------------------------------------------------------------
Public Function IsLegalFileName(ByVal str As String) As Boolean
   If Not (str Like "*[/\:*?""<>|]*") Then
     IsLegalFileName = True
   End If
 End Function
'-----------------------------------------------------------------------------
'Source: Adapted from Ron deBruin @ http://www.rondebruin.nl/mail/folder2/files.htm
Sub processEmail(OutApp As Object, emailTo As String, emailSubject As String, emailBody As String, emailFile As String)
'Working in 2000-2010
    'Dim outMail As Outlook.MailItem 'early binding
    Dim outMail As Object 'late binding

    Application.StatusBar = "Sending Email To: " & emailTo & " File: " & emailFile
       
    'Set outMail = OutApp.CreateItem(olMailItem) 'early binding
    Set outMail = OutApp.CreateItem(0) 'late binding

    With outMail
        .To = emailTo
        .Subject = emailSubject
        .body = emailBody
        .attachments.Add emailFile
        '.Display
        .Send
    End With
    
    Set outMail = Nothing
    
    Application.StatusBar = False
End Sub

