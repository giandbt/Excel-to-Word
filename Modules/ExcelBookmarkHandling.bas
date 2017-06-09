Attribute VB_Name = "ExcelBookmarkHandling"
Option Explicit
Public Const RANGE_OBJ = 1
Public Const RANGE_NAME = 2
Public Const SHAPE_OBJ = 3
Public Const CHART_OBJ = 4
Public Const CHART_EMB = 5
Public myBM As BM_Indicators
Public Sub updateWordFromExcel(Optional control As Object) 'IRibbonControl
Dim validError As String
Dim strNameScope As String
Dim xMsg As Long
Dim strPathFName As String
Dim wkb As Workbook
Dim wks As Worksheet
'Dim oWA As Word.Application 'early binding
Dim oWA As Object 'late binding
'Dim oWD As Word.Document 'early binding
Dim oWD As Object 'late binding
'Dim bkMk As Word.Bookmark 'early binding
Dim BkMk As Object 'late binding
Dim fPath As String
Dim fName2 As String
Dim PDFname As String
Dim PDFname2 As String
Dim fRange As Range
Dim tbl As Object
Dim rw As Object
Dim cl As Object
Dim dataRow As Boolean
'Dim FSO As FileSystemObject 'early binding
Dim FSO As Object 'late binding
Dim BM_col As New BM_Indicators 'collection of bookmark indicators in Excel workbook
Dim eMail_Col As New BM_Indicators ' use same container for email address
Dim myObj As BM_Indicator
Dim bMultiCellOShape As Boolean
Dim bPasteChartSheet As Boolean
Dim bPasteChartEmbed As Boolean
Dim myObjCopy As Object
Dim bResult As Boolean
Dim i As Long
Dim lLoop As Long
Dim rIncrement As Range
Dim lStart As Long
Dim lEnd As Long
Dim xCalc As Long
Dim bDraftPreview As Boolean
Dim bPasteEnhMeta As Boolean
Dim fileAttach As String
'Dim OutApp As Outlook.Application 'early binding
Dim OutApp As Object 'late binding

    If Application.Workbooks.Count = 0 Then
        MsgBox "No files open to process"
        Exit Sub
    End If
    
    If ActiveSheet.Type <> xlWorksheet Then
        MsgBox "You can only run ExcelToWord! functions from Excel Worksheets (e.g., Not from Chart Sheets, etc.)", vbCritical
        Exit Sub
    End If
    
    xCalc = Application.Calculation
    
    Application.StatusBar = "Update Word From Excel: Initialization..."
    
'The Configuration Options panel should not have saved a set of invalid options, but to be sure,
'complete a final pass of run-through validations prior to the update.  Recall, it could be days, weeks, or months since this workbook
'was originally created and successfully completed an ExcelToWord! update.  As a result, file paths, templates, etc., could have been
'deleted, renamed, or relocated...

'Checking all relevant options
    
    If myEvaluate(CONFIG_SCOPE) = "" Or (myEvaluate(CONFIG_SCOPE) = "Worksheet" And _
        myEvaluate(CONFIG_SHEET) = "") Then 'scope has not been defined, go to Configurator
        
        xMsg = MsgBox("Configurator settings have not been defined.  Proceed to Configuration Options?", vbYesNo, "Proceed to Configuration Options?")
        If xMsg = vbYes Then
            GoTo backToUserform
        Else
            GoTo gracefulExit
        End If
    End If
    'first, validate all entries in the current configuration (as source files may have been deleted/renamed since the configuration was set up.
    Call setPublicVariables 'load configuration for current activity
        
    'check scope
    strNameScope = myEvaluate(CONFIG_SCOPE)
    If strNameScope = "" Then
        validError = "CONFIG_SCOPE ERROR:  Please revisit the Configuration Options panel, as there's some confusion about the scope.  " & _
            "No value for scope (Worksheet or Workbook)"
        GoTo backToUserform
    End If
        
    'ensure word template exists - the one that should have been generated
    If strWD_TemplOpt <> "OWN" Then
        If strWD_TemplateBMFile = vbNullString Or Not validateFileFolderSelection(strWD_TemplFile, "Word", "template", False) Then
            validError = "Word Template File ERROR:  The path\filename no longer exists, or needs to be re-generated" & vbCrLf & vbCrLf & "[path\filename]: " & strWD_TemplFile & vbCrLf & vbCrLf & "You may need to just Generate Word Bookmarks, or ..."
            GoTo backToUserform
        End If
    Else
        strWD_TemplateBMFile = strWD_TemplFile 'OWN option does not require BM File generation, but name it now, as the rest of the code depends on it
    End If
    

    'notify user with options if word document filename exists at that path - overwrite or cancel
    If bAftUpdSave Then
        'ensure word document path still exists
        If strWD_DocPath = vbNullString Or Not validateFileFolderSelection(strWD_DocPath, "Word", "document", True) Then
            validError = "New Word Document Path ERROR:  The path\filename no longer exists" & vbCrLf & vbCrLf & "[path\filename]: " & strWD_DocPath
            GoTo backToUserform
        ElseIf strWD_DocFile = vbNullString Then
            validError = "New Word Document File ERROR:  The filename chosen is no longer valid.  You might try save/close Excel, then reload your workbook and check Configuration Options"
            GoTo backToUserform
        End If
    End If
    
    'open word template as a document
    'Set FSO = New FileSystemObject 'early binding
    Set FSO = CreateObject("Scripting.FileSystemObject") 'late binding
    
    Set wkb = ActiveWorkbook
    Set wks = wkb.ActiveSheet
    
    fPath = getPathFromPathFName(strWD_TemplateBMFile)
    If bAftUpdPDF Then 'get path for PDF file generation & advise user
        If bAftUpdSave Then
            PDFname = strWD_DocPath & "\" & strWD_DocFile & ".pdf"
            MsgBox "PDF File will be saved in directory:" & vbCrLf & vbCrLf & strWD_DocPath & vbCrLf & vbCrLf & "The same as the generated Word Document", vbOKOnly
        Else
            PDFname = Left(strWD_TemplateBMFile, InStr(strWD_TemplateBMFile, ".") - 1) & ".pdf"
            MsgBox "PDF file will be saved in directory:" & vbCrLf & vbCrLf & fPath & vbCrLf & vbCrLf & "The same as the existing Word Template", vbOKOnly
        End If
    End If
    
    If FSO.fileExists(strWD_TemplateBMFile) Then
        
        'start new instance of Word, regardless if an instance exists
        'Set oWA = New Word.Application 'early binding
        Set oWA = CreateObject("Word.Application")
        
        'Prepare for Increment generation
        If bXL_Increment Then
            lStart = Range(strXL_RefStart).Value
            lEnd = Range(strXL_RefEnd).Value
        Else
            lStart = 1
            lEnd = 1
        End If
        
        For lLoop = 0 To lEnd - lStart
        
            If bXL_Increment Then 'set Incrementer value so data refresh is forced
                Range(strXL_RefCounter).Value = lStart + lLoop
                If xCalc = xlCalculationManual Then Application.Calculate
            End If
            
            Set oWD = oWA.Documents.Add(Template:=strWD_TemplateBMFile) 'Create New Document From Template
            oWA.Visible = oWA_VISIBLE
            
            'traverse all bookmarks and ensure that those bookmarks exist in Excel, looking at selected options - range, labels, or both
            For Each BkMk In oWD.Bookmarks 'first pass to build collection of Excel bookmark indicator (objects) associated with each Word bookmark
                'find corresponding Excel key that matches bookmark
                'look in range names first, then shape names (e.g., charts,images, etc.)
                'then bookmark indicators, as prescribed by the Configuration options selected
    
                Application.StatusBar = "[" & lLoop + 1 & "]:" & "Testing for Bookmark: " & BkMk.Name & "..."
                
                'search range names, then shape names option
                Select Case strXL_TemplOpt:
                    Case "RANGE":  'search range names, then shape names for bookmark indicators
                        bResult = searchRangeShapes(BM_col, BkMk, bXL_SpanWorkbook)
                        
                    Case "RANGE_AND_CELL": 'search range names, then shape names, then CELLS for bookmark indicators
                        bResult = searchRangeShapes(BM_col, BkMk, bXL_SpanWorkbook)
                        If Not bResult Then 'if not found in range, then look at CELL level
                            bResult = searchCells(BM_col, BkMk.Name, bXL_SpanWorkbook)
                        End If
                        
                    Case "CELL": 'search CELLS for bookmark indicators
                        bResult = searchCells(BM_col, BkMk.Name, bXL_SpanWorkbook)
                End Select
                
                If Not bResult Then 'bookmark not found!
                    xMsg = MsgBox("Cannot Find Excel data for bookmark: " & BkMk.Name & ".  Continue anyway?", vbOKCancel, "Hit OK to Continue, Cancel to Abort")
                    If xMsg = vbCancel Then GoTo gracefulExit
                End If
            
            Next BkMk
                               
            'now search for eMail marker in workbook [[eMail]]
            If strAftUpdEmail <> "" Then
                bResult = searchCells(eMail_Col, "eMailTo", bXL_SpanWorkbook) 'just add the eMail indicator to the bookmark indicators collection
                If bResult Then
                    bResult = searchCells(eMail_Col, "emailSubject", bXL_SpanWorkbook)
                    If bResult Then
                        bResult = searchCells(eMail_Col, "emailBody", bXL_SpanWorkbook)
                    End If
                End If
                
                If Not bResult Then 'bookmark not found!
                    xMsg = MsgBox("Cannot Find Excel data for eMail address: [[eMailTo]], [[eMailSubject]], or [[eMailBody]] is missing. Continue anyway?", vbOKCancel, "Hit OK to Continue, Cancel to Abort")
                    If xMsg = vbCancel Then GoTo gracefulExit
                End If
                
                On Error Resume Next
                Set OutApp = GetObject(, "Outlook.Application")
                If OutApp Is Nothing Then
                    'Set OutApp = New Outlook.Application 'early binding
                    Set OutApp = CreateObject("Outlook.Application") 'late binding
                End If
                On Error GoTo 0
            End If
            
            'now loop through collection of found bookmark indicators, and output results to Word template
            For Each BkMk In oWD.Bookmarks 'second pass:  now we have matching Excel bookmark indicators and Word objects
            
                Application.StatusBar = "[" & lLoop + 1 & "]:" & "Second Pass:  Updating Word bookmarks from Excel for Bookmark: " & BkMk.Name & "..."
                
                bMultiCellOShape = False
                bPasteChartSheet = False
                bPasteChartEmbed = False
                
                On Error Resume Next 'recall, user may have allowed "Continue anyway" if bookmark indicator wasn't found
                Set myObj = BM_col(BkMk.Name)
                If Err.Number <> 0 Then 'assumed missed bookmark, but continue
                    'do nothing
                    On Error GoTo 0
                ElseIf Not myObj Is Nothing Then
                    On Error GoTo 0
                    
                    'determine if type resolves to a single cell, a range > 1 cell, or a shape
                    Select Case myObj.BM_Type
                        Case RANGE_NAME:
                            bMultiCellOShape = IIf(myObj.obj.RefersToRange.Count > 1, True, False)
                            Set myObjCopy = myObj.obj.RefersToRange
                        Case RANGE_OBJ:
                            bMultiCellOShape = False
                            Set myObjCopy = myObj.obj
                        Case SHAPE_OBJ:
                            bMultiCellOShape = True
                            Set myObjCopy = myObj.obj
                        Case CHART_OBJ:
                            Set myObjCopy = myObj.obj.ChartArea
                            bPasteChartSheet = True
                        Case CHART_EMB:
                            Set myObjCopy = myObj.obj
                            bPasteChartEmbed = True
                    End Select
                    
                    If bPasteChartSheet Or bPasteChartEmbed Then
                        'need to test if the bookmark in Word is a Shape, or Text
                        Dim r As Object
                        Set r = oWA.Selection.GoTo(what:=wdGoToBookmark, Name:=BkMk.Name)
                        If r.Text <> "" Then 'the bookmark is referencing text - a normal text-based bookmark indicator
                            myObjCopy.Copy
                            On Error Resume Next
                            BkMk.Range.PasteSpecial Link:=True, Placement:=wdInLine, DataType:=iXL_TemplOptShapePaste
                            If Err.Number <> 0 Then
                                BkMk.Range.PasteSpecial Link:=True, Placement:=wdInLine, DataType:=wdPasteEnhancedMetafile
                                bPasteEnhMeta = True
                            End If
                            On Error GoTo 0
                            Application.CutCopyMode = False
                        ElseIf Not pastePicToBkMk(oWA, myObjCopy, BkMk) Then 'the bookmark is referencing a Shape, so paste via fill effects of the Shape
                            'paste shape/image/chart as picture into Word Shape bookmark
                            xMsg = MsgBox("Could not paste shape/image as a fill picture for bookmark: " & BkMk.Name & "." & _
                                vbCrLf & vbCrLf & "Continue anyway?", vbYesNo, "Hit YES to Continue, NO to Abort")
                            If xMsg = vbNo Then GoTo gracefulExit
                        End If
                        
                    ElseIf bMultiCellOShape Then
                        myObjCopy.Copy
                        On Error Resume Next
                            BkMk.Range.PasteSpecial Link:=True, Placement:=wdInLine, DataType:=iXL_TemplOptShapePaste
                            If Err.Number <> 0 Then
                                BkMk.Range.PasteSpecial Link:=True, Placement:=wdInLine, DataType:=wdPasteEnhancedMetafile
                                bPasteEnhMeta = True
                            End If
                            On Error GoTo 0
                        Application.CutCopyMode = False
                    Else
                        myObjCopy.Copy
                        If myObjCopy.Value <> "" Then
                            BkMk.Range.PasteSpecial Link:=True, Placement:=wdInLine, DataType:=1
                        Else
                            BkMk.Range.Text = myObjCopy.Value 'use base format for all else
                        End If
                        Application.CutCopyMode = False
                    End If
        
                End If
                On Error GoTo 0
            Next BkMk
            
            'The following code assumes that the application requires a list of items which can vary from 1 to unlimited
            If bWD_Table Then
                'So, there are 1 to many rows of BookMarks - e.g., invoice lineItems, For Example:
                'lineItem1, description1, amount1
                'lineItem2, description2, amount2
                '...
                'lineItem-n, description-n, amount-n
                '
                'As a result, if the Excel template uses only the first few line items, the remaining line items would be a blank
                'copy from Excel to Word, leaving blank lines in the Word Template - and perhaps an unattractive gap between a list of line items,
                'and the rest of the invoice.
                '
                'The following loop traverses all tables in the template and deletes lineItems that are blank
                
                Application.StatusBar = "[" & lLoop + 1 & "]:" & "Cleaning Word Template Tables..."
                
                'If there are any tables in the Word template, and their row is empty, then delete that empty row
                For Each tbl In oWD.Tables
                    For Each rw In tbl.Rows 'examine each row
                        dataRow = False
                        For Each cl In rw.Cells 'look at all cells in each row
                            If Len(Trim(Application.WorksheetFunction.Clean(cl.Range.Text))) > 0 Then
                                dataRow = True 'if there's data in any cell, then there's data in the row
                                Exit For
                            End If
                        Next cl
                        If Not dataRow Then
                            rw.Delete 'delete any rows in the table that all cells on that row are empty
                        End If
                    Next rw
                Next tbl
            End If
            
            'The document is now complete, all that remains is to print, extract to PDF, and/or save, then close the file, per Configuration Options
            If bAftUpdPrint Then
                Application.StatusBar = "[" & lLoop + 1 & "]:" & "Printing Word Document..."
                oWD.PrintOut
            End If
                
            If bAftUpdPDF Then
                'Save Word Document as PDF
                'for Office 2007 with Office PDF Add-On from http://labnol.blogspot.com/2006/09/office-2007-save-as-pdf-download.html, or
                'http://www.ehow.com/how_7184784_save-word-docs-pdf-vba.html
                
                If bXL_Increment Then
                    PDFname2 = Left(PDFname, Len(PDFname) - 4) & "_" & Format(lLoop + 1, "000") & ".pdf"
                End If
                
                Application.StatusBar = "[" & lLoop + 1 & "]:" & "Generating PDF file: " & PDFname2
                
                On Error Resume Next
                oWD.ExportAsFixedFormat OutputFileName:=PDFname2, ExportFormat:= _
                    wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
                    wdExportOptimizeForPrint, Range:=wdExportAllDocument, _
                    Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
                    CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
                    BitmapMissingFonts:=True, UseISO19005_1:=False
                If Err.Number <> 0 Then
                    MsgBox "Unable to SaveAs/ExportTo PDF - you are either: " & vbCrLf & vbCrLf & _
                        "1) Running Excel 2003 or earlier, " & vbCrLf & _
                        "2) Running Excel 2007 without the required Office 2007 Save as PDF Add-on (See http://www.microsoft.com/download/en/details.aspx?id=7)" & vbCrLf & _
                        " or " & vbCrLf & _
                        "3) There's a problem with your Save as PDF capability in either Excel 2007 or Excel 2010." & vbCrLf & vbCrLf & _
                        "Please repair and try again", vbCritical, "Skipping Save as PDF step..."
                End If
                On Error GoTo 0
            End If
                
            If bAftUpdSave Then
                'Save Word document, in current format (e.g., doc, docx, etc.) then close file
                
                If bXL_Increment Then
                    fName2 = strWD_DocFile & "_" & Format(lLoop + 1, "000")
                Else
                    fName2 = strWD_DocFile
                End If
                    
                Application.StatusBar = "[" & lLoop + 1 & "]:" & "Saving Word Document: " & strWD_DocPath & "\" & fName2
                oWD.SaveAs Filename:=strWD_DocPath & "\" & fName2
                oWD.Close
                Set oWD = Nothing
            ElseIf bAftUpdDelete Then 'otherwise, done with file, without save
                oWD.Close SaveChanges:=False
            Else 'then just preview the new Word document
                oWA.Visible = True
                bDraftPreview = True
                MsgBox "Toggle to Word document for Preview", vbOKOnly, "Terminating operation after 1st draft generated"
                GoTo gracefulExit
            End If
            
            If strAftUpdEmail <> "" And Not eMail_Col Is Nothing Then
                'eMail the PDF or Word document
                If UCase(strAftUpdEmail) = UCase("ePDF") Then 'process email w/ PDF
                    fileAttach = PDFname2
                Else 'process email w/ Word document
                    fileAttach = oWD.Name
                End If
                
                If fileAttach <> "" Then
                    Call processEmail(OutApp, eMail_Col("emailTo").obj, eMail_Col("emailSubject").obj, eMail_Col("emailBody").obj, fileAttach)
                End If
            End If
            
            'clean up before next pass
            BM_col.RemoveAll
            Set BM_col = Nothing
            If Not eMail_Col Is Nothing Then 'prepare for next eMail address, if we're emailing
                eMail_Col.RemoveAll
                Set eMail_Col = Nothing
            End If
        Next lLoop
        
        Application.StatusBar = False
        MsgBox "Successful ExcelToWord! production process", vbOKOnly
       
    Else
        MsgBox "Template file: " & strWD_TemplateBMFile & " not found at " & fPath & " - please create Template and try again", vbCritical, "Aborting"
    End If

    GoTo gracefulExit
    
backToUserform:
    If validError <> "" Then
        xMsg = MsgBox(validError & vbCrLf & vbCrLf & "Open Configuration Options to make changes?", vbYesNo, _
            "Configurator Error: Hit YES to pull up Configuration Options, NO to Abort")
        If xMsg = vbYes Then Call showConfigurator
    Else
        Call showConfigurator
    End If
    
gracefulExit:
    Application.StatusBar = False
    
    If Not bDraftPreview Then 'only if successful generation of draft will this be skipped
        'clean up open word document and application, if any
        If Not oWD Is Nothing Then oWD.Close SaveChanges:=False
        If Not oWA Is Nothing Then oWA.Quit
    End If
    
    BM_col.RemoveAll
    Set BM_col = Nothing
    
    If bPasteEnhMeta Then MsgBox "Could not paste all objects according to style selected, so pasted as Enhanced Metafile, instead"
End Sub

'Private Function searchRangeShapes(BM_col As BM_Indicators, bkMk As Word.Bookmark, bXL_SpanWorkbook As Boolean) As Boolean 'early binding
Private Function searchRangeShapes(BM_col As BM_Indicators, BkMk As Object, bXL_SpanWorkbook As Boolean) As Boolean 'late binding
Dim wkb As Workbook
Dim wks As Worksheet
Dim cht As Chart
Dim myActWks As Worksheet
Dim rName As Name
Dim shp As Shape
Dim strSearch As String
Dim xMsg As Long
Dim myQuote_char As String
        
    Set wkb = ActiveWorkbook
    Set myActWks = wkb.ActiveSheet
    
    'Search for Range name matching Excel Bookmark Indicator name, at ActiveSheet level, then Workbook level, exiting on first instance found
    If Not bXL_SpanWorkbook Then 'search within ActiveSheet scope, only
        If InStr(myActWks.Name, SPACE_CHAR) <> 0 Then
            myQuote_char = QUOTE_CHAR
        Else
            myQuote_char = vbNullString
        End If
        strSearch = UCase(myQuote_char & myActWks.Name & myQuote_char & "!" & BkMk.Name)
        
        On Error Resume Next
        Set rName = myActWks.Names(strSearch)
        If Err.Number = 0 Then
            BM_col.Add BkMk.Name, rName, RANGE_NAME
            searchRangeShapes = True
            Exit Function 'stop when first instance is found
        End If
        On Error GoTo 0
    Else
        
        On Error Resume Next
        Set rName = wkb.Names(BkMk.Name)
        If Err.Number = 0 Then
            BM_col.Add BkMk.Name, rName, RANGE_NAME
            searchRangeShapes = True
            Exit Function 'stop when first instance is found
        End If
        On Error GoTo 0
        
        'finally, find first range name that matches at the worksheet level - span workbook has workbook level name priority,
        'then worksheet name, starting with activesheet as priority
        
        'Check ActiveSheet
        If InStr(myActWks.Name, SPACE_CHAR) <> 0 Then
            myQuote_char = QUOTE_CHAR
        Else
            myQuote_char = vbNullString
        End If
        strSearch = UCase(myQuote_char & myActWks.Name & myQuote_char & "!" & BkMk.Name)
        
        On Error Resume Next
        Set rName = myActWks.Names(strSearch)
        If Err.Number = 0 Then
            BM_col.Add BkMk.Name, rName, RANGE_NAME
            searchRangeShapes = True
            Exit Function 'stop when first instance is found
        End If
        On Error GoTo 0
        
        'now check the rest of the sheets
        For Each wks In wkb.Worksheets
            If wks.Name <> myActWks.Name Then
                If InStr(wks.Name, SPACE_CHAR) <> 0 Then
                    myQuote_char = QUOTE_CHAR
                Else
                    myQuote_char = vbNullString
                End If
                strSearch = UCase(myQuote_char & wks.Name & myQuote_char & "!" & BkMk.Name)
                                
                On Error Resume Next
                Set rName = wks.Names(strSearch)
                If Err.Number = 0 Then
                    BM_col.Add BkMk.Name, rName, RANGE_NAME
                    searchRangeShapes = True
                    Exit Function 'stop when first instance is found
                End If
                On Error GoTo 0
            End If
        Next wks
    End If
    
    'if we didn't find it in a Range, then let's look at shapes - e.g., charts, images, and other assorted shapes, using the Shapes collection
    'search workbook_level names, then worksheet names, on every sheet, until found
    If Not bXL_SpanWorkbook Then
        On Error Resume Next
        Set shp = myActWks.Shapes(BkMk.Name)
        If Err.Number = 0 Then
            If shp.Type = msoChart Then 'embedded chart
                BM_col.Add BkMk.Name, shp, CHART_EMB
            Else
                BM_col.Add BkMk.Name, shp, SHAPE_OBJ
            End If
            searchRangeShapes = True
            Exit Function 'stop when first instance is found
        End If
        On Error GoTo 0
        
        'Chart sheets can exist, even though bXL_SpanWorkbook is false, so test for those
        On Error Resume Next
        Set cht = wkb.Charts(BkMk.Name)
        If Err.Number = 0 Then
            BM_col.Add BkMk.Name, cht, CHART_OBJ
            searchRangeShapes = True
            Exit Function
        End If
        On Error GoTo 0
    Else    'search workbook_level shape names, then worksheet shape names, on every sheet
            'check for chart sheet, first
            
        On Error Resume Next
        Set cht = wkb.Charts(BkMk.Name)
        If Err.Number = 0 Then
            BM_col.Add BkMk.Name, cht, CHART_OBJ
            searchRangeShapes = True
            Exit Function
        End If
        
        'then look at embedded shapes at the worksheet level
        For Each wks In wkb.Worksheets
            On Error Resume Next
            Set shp = wks.Shapes(BkMk.Name)
            If Err.Number = 0 Then
                If shp.Type = msoChart Then 'embedded chart
                    BM_col.Add BkMk.Name, shp, CHART_EMB
                Else
                    BM_col.Add BkMk.Name, shp, SHAPE_OBJ
                End If
                searchRangeShapes = True
                Exit Function 'stop when first instance is found
            End If

        Next wks
    End If
    
    'otherwise, fail out
End Function
Private Function searchCells(BM_col As BM_Indicators, strBkMk As String, bXL_SpanWorkbook As Boolean) As Boolean
Dim fRange As Range
Dim wkb As Workbook
Dim wks As Worksheet
Dim myActWks As Worksheet
Dim focusRange As Range

    'routine searches for Excel bookmark indicators, identifying each corresponding data-point adjacent to the indicator inside the BM_Indicators class collection
    
    Set wkb = ActiveWorkbook
    Set myActWks = wkb.ActiveSheet
    
    For Each wks In wkb.Worksheets
        If bXL_SpanWorkbook Or (bXL_SpanWorkbook = False And wks.Name = myActWks.Name) Then 'search all worksheets, or active sheet
            Set fRange = wks.Cells.Find(what:="[[" & strBkMk & "]]", LookIn:=xlValues, lookat:=xlWhole)
            If Not fRange Is Nothing Then
                
                On Error Resume Next
                
                Select Case strXL_TemplOptCell
                    Case "Left": Set focusRange = fRange.Offset(0, 1)
                    Case "Above": Set focusRange = fRange.Offset(1, 0)
                    Case "Right": Set focusRange = fRange.Offset(0, -1)
                    Case "Below": Set focusRange = fRange.Offset(-1, 0)
                End Select
                
                If Err.Number <> 0 Then
                    MsgBox "You indicated bookmark indicators would be adjacent " & UCase(strXL_TemplOptCell) & " of the data, while bookmark indicator " & strBkMk & " at " & "'" & fRange.Worksheet.Name & "'!" & fRange.Address & " throws an error when that offset is performed." & vbCrLf & vbCrLf & "Please recast bookmark: " & strBkMk, vbCritical, "Aborting..."
                    searchCells = False
                    Exit Function
                End If
                
                On Error GoTo 0
                
                BM_col.Add strBkMk, focusRange, RANGE_OBJ
                searchCells = True
                Exit Function 'stop when first instance is found
            End If
        End If
    Next wks

End Function
'Private Function pastePicToBkMk(oWA As Word.Application, myObjCopy As Object, bkMk As Word.Bookmark) As Boolean 'early binding
Private Function pastePicToBkMk(oWA As Object, myObjCopy As Object, BkMk As Object) As Boolean 'late binding
Dim strTmpPicFile As String
Dim r As Object

'logic to change bookmark shape fill effects, importing temporary image

    On Error GoTo errHandler
    
    'first, save the image to a temporary file
    strTmpPicFile = export(myObjCopy)

    'then, navigate to the bookmark, and change the fill effects, importing the image
    Set r = oWA.Selection.GoTo(what:=wdGoToBookmark, Name:=BkMk.Name)
    
    'no line around shape and ensure picture fits re: aspect ratio
    r.ShapeRange.Fill.Transparency = 0#
    r.ShapeRange.Line.Visible = msoFalse
    r.ShapeRange.LockAspectRatio = msoFalse
    
    'replace recorded filename with temporary file name just generated
    r.ShapeRange.Fill.UserPicture strTmpPicFile
        
    pastePicToBkMk = True
    
    GoTo gracefulExit
    
errHandler:
    pastePicToBkMk = False
    
gracefulExit:
    On Error Resume Next
    Kill strTmpPicFile 'delete temporary file
    On Error GoTo 0
End Function
