Attribute VB_Name = "WordBookmarkHandling"
Option Explicit

Public Sub generateWordBookmarks(Optional control As Object) 'IRibbonControl
Dim xMsg As Long
Dim myMsg As String

'The Configuration Options panel should not have saved a set of invalid options, but to be sure,
'complete a final pass of run-through validations prior to the update.  Recall, it could be days, weeks, or months since this workbook
'was originally created and successfully completed an ExcelToWord! update.  As a result, file paths, templates, etc., could have been
'deleted, renamed, or relocated...

    If Application.Workbooks.Count = 0 Then
        MsgBox "No files open to process"
        Exit Sub
    End If
    
    If ActiveSheet.Type <> xlWorksheet Then
        MsgBox "You can only run ExcelToWord! functions from Excel Worksheets (e.g., Not from Chart Sheets, etc.)", vbCritical
    
    ElseIf myEvaluate(CONFIG_SCOPE) = "" Or (myEvaluate(CONFIG_SCOPE) = "Worksheet" And _
        myEvaluate(CONFIG_SHEET) = "") Then 'scope has not been defined, go to Configurator
        
        xMsg = MsgBox("Configurator settings have not been defined.  Proceed to Configuration Options?", vbYesNo, "Proceed to Configuration Options?")
        If xMsg = vbYes Then Call showConfigurator
    Else
        'first, validate all entries in the current configuration (as source files may have been deleted/renamed since the configuration was set up.
        Call setPublicVariables 'load configuration for current activity
        
        'check for Word template existence
        If Not validateFileFolderSelection(strWD_TemplFile, "Word", "template", False) Then
            MsgBox "The path\filename no longer exists" & Chr(10) & Chr(10) & strWD_TemplFile & Chr(10) & Chr(10) & "Please return to Configuration Options and Fix entry, or delete entry and BROWSE for file", vbOKOnly, "Configurator Error"
        ElseIf strWD_TemplOpt = "OWN" Then
            MsgBox "Configuration Options set to ""OWN"" therefore cancelling request to generate bookmarks.  Instead, you may proceed directly to the Update Word from Excel process."
        Else
            Call readWordDocMakeBookmarks(IIf(strWD_TemplOpt = "GENERIC", True, False), strWD_TemplFile)
        End If
    End If
End Sub
Private Sub readWordDocMakeBookmarks(bGeneric As Boolean, fPathFname As String)
'Dim oWA As Word.Application 'early binding
Dim oWA As Object 'late binding
'Dim oWD As Word.Document 'early binding
Dim oWD As Object 'late binding
'Dim para As Paragraph 'early binding
Dim para As Object
Dim bmks As Variant
Dim i As Integer
'Dim myDict As Scripting.Dictionary 'early binding
Dim myDict As Object 'late binding
Dim cntDict As Long
Dim regExPattern As String
Dim bResult As Boolean
Dim fName As String, fPath As String, fBMName As String
Dim fNameExt As String
Dim tempBMK As String
Dim objWkbSht As Object

'Rules for Bookmarks - NO duplicates, NO spaces.  Must start with [[  and end with ]], may include alphanumeric and underscore only
'This app will find proposed bookmarks in word document, and make them according to the book mark "name" inside the [[name]] brackets
'It will then save the file as a NEW TEMPLATE to be used with this application, named template_BM.dotx
'On the active sheet of the active workbook will be a range name called "WordDoc" that will be the name of the Word template
'to be found in the active workbook's path.

'If bookmarks already exist in the document, the new bookmark will overwrite the old.  Formfields having same name as proposed bookmarks will prompt
'option to skip that bookmark (encouraging user to clean up, after) or abort the update process.

    'start new instance of Word, regardless if an instance exists
    'Set oWA = New Word.Application 'early binding
    Set oWA = CreateObject("Word.Application") 'late binding
    
    'Set myDict = New Scripting.Dictionary 'early binding
    Set myDict = CreateObject("Scripting.Dictionary") 'late binding
    
    fPath = getPathFromPathFName(fPathFname)
    fName = Right(fPathFname, Len(fPathFname) - Len(fPath))
    
    fNameExt = getFileExt(fName) 'get file extension
    
    fBMName = Left(fName, Len(fName) - Len(fNameExt)) & "_BM" & fNameExt
    
    Set oWD = oWA.Documents.Open(Filename:=fPath & fName, ReadOnly:=True, AddToRecentFiles:=False) 'ReadOnly - never subject original template to corruption, .Add opens document based on template, .Open opens the Word TEMPLATE

    oWA.Visible = oWA_VISIBLE
    
    regExPattern = "\[{2}[A-Za-z0-9_]+\]{2}" 'looks for strings like [[alphanumeric or underscore]] spaces in BM's not permitted, also no duplicates
    
    For Each para In oWD.Paragraphs
        bmks = RegExpFind(para.Range.Text, regExPattern)
               
        On Error GoTo flagError
        
        If Not IsNull(bmks) Then
            For i = 0 To UBound(bmks)
                Application.StatusBar = "Processing bookmark " & bmks(i) & "..."
                cntDict = cntDict + 1 'new bookmark counter
                
                'do some validation - ensure GENERIC bookmarks all are of the type [[BM]], and that INTELLIGENT/PERSONALIZED bookmarks are unique via dictionary
                If bGeneric And bmks(i) <> "[[BM]]" Then Err.Raise 3, Description:="GENERIC bookmark is invalid - must be EXACTLY ""[[BM]]""" & _
                    Chr(10) & "BookMark: " & bmks(i) & Chr(10) & "Paragraph: " & para.Range.Text
                
                If bGeneric Then
                    tempBMK = Left(bmks(i), Len(bmks(i)) - 2) & "_" & cntDict & "]]" 'embed counter in bookmark name
                Else
                    tempBMK = bmks(i)
                End If
                
                'continue validation - ensure bookmark is unique, and if so, then generate bookmark
                If Not myDict.Exists(tempBMK) Then
                    myDict.Add tempBMK, cntDict
                    
                    'now, modify the Word Template, setting the bookmark
                    bResult = setWordBookMark(oWD, para, tempBMK, bGeneric)
                    If Not bResult Then Err.Raise 2, Description:="Cannot create bookmark in Word for some reason" & Chr(10) & "BookMark: " & bmks(i) & Chr(10) & "Paragraph: " & para.Range.Text
                Else
                    Err.Raise 1, Description:="Error:  Duplicate found on proposed bookmark " & tempBMK & ": Bookmark proposed does not follow rules: " _
                        & Chr(10) & Chr(10) & "Rules for Bookmarks - NO duplicates, NO spaces.  Must start with [[  and end with ]]," & _
                        " may include alphanumeric and underscore only"
                End If
            Next i
        End If
        On Error GoTo 0
    Next para
    
    Application.StatusBar = "Saving Bookmark Template: " & fPath & fBMName & "..."
    
    'Note - FileFormat:= not needed - save in same format
    oWD.SaveAs Filename:=fPath & fBMName, _
        LockComments:=False, Password:="", AddToRecentFiles:=True, _
        WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
        
    'now, add this to the set of options for this sheet/workbook, for retrieval on the ExcelToWord! Update process
    strWD_TemplateBMFile = fPath & fBMName
    
    Set objWkbSht = IIf(bXL_SpanWorkbook, ActiveWorkbook, ActiveSheet)
    
    objWkbSht.Names.Add Name:="ETW_strWD_TemplateBMFile", RefersTo:=strWD_TemplateBMFile, Visible:=NAME_VISIBLE
    
    Application.StatusBar = False
    MsgBox "Successful Creation of " & myDict.Count & " Bookmarks" & Chr(10) & Chr(10) & "Revised Template File Has Been Saved: " & fBMName


gracefulExit:

    Application.StatusBar = False
    myDict.RemoveAll
    Set myDict = Nothing
    oWA.Quit
    
    Exit Sub
    
flagError:
    If Err.Number < 5 Then
        MsgBox "Error: " & Err.Number & "->" & Err.Description & Chr(10) & "Please correct problem with template/workbook and try again", vbCritical, "Aborting!..."
    Else
        MsgBox "VBA Error: " & Err.Number & "->" & Err.Description & Chr(10) & "Hit ok to enter Debugger", vbOKOnly, "Please correct VBA code - Aborting"
        Stop 'hit F8 to resume at error line for debug mode
        Resume
    End If
    
    Resume gracefulExit
    
End Sub
'Private Function setWordBookMark(oWD As Word.Document, para As Word.Paragraph, bmStr As Variant, bGeneric As Boolean) As Boolean 'early binding
Private Function setWordBookMark(oWD As Object, para As Object, bmStr As Variant, bGeneric As Boolean) As Boolean 'late binding
'Dim oWA As Word.Application 'early binding
Dim oWA As Object 'late binding
'Dim oBMK As Word.Bookmark 'early binding
Dim oBMK As Object 'late binding
Dim BM_Name As String
Dim xMsg As Long
Dim bDelete As Boolean

'Searches for Word bookmark indicators, then creates a bookmark for each.
'Generic bookmark indicators are incremented and "flagged" (e.g., [[BM_XX]]) with numeric increments, in the text of the template, as well.

    bDelete = True
    
    BM_Name = Left(Right(bmStr, Len(bmStr) - 2), Len(Right(bmStr, Len(bmStr) - 2)) - 2) 'eliminate the left and right [[ ]] braces from BookMark name

    Set oWA = oWD.Parent
    
    oWA.Selection.Find.ClearFormatting

    With oWA.Selection.Find
    If bGeneric Then
        .Text = "[[BM]]"
        .Replacement.Text = bmStr
    Else
        .Text = bmStr
    End If
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    If bGeneric Then
        oWA.Selection.Find.Execute Replace:=wdReplaceOne
    Else
        oWA.Selection.Find.Execute
    End If
    
    If BookmarkExists(oWD, BM_Name) Then 'existing bookmarks will be overwritten, but test formfields, first
        Set oBMK = oWD.Bookmarks(BM_Name)
        
        If ISFormfield(oBMK) Then
            xMsg = MsgBox("Bookmark: " & BM_Name & " already exists as a Form Field - do you want to SKIP this bookmark (YES - SKIP, keeping the bookmark/formfield ""as-is"" (note, you'll want to eliminate or restate a new name for the [[" & BM_Name & "]] in the Word template),CANCEL - Abort the process?", vbYesNoCancel, "YES - Skip & Continue, CANCEL - Abort")
            If xMsg = vbYes Then
                setWordBookMark = True
                Exit Function
            Else
                setWordBookMark = False
                Exit Function
            End If
        End If
        
        oBMK.Delete
        
    End If
    
    'now, create the bookmark
    With oWD.Bookmarks 'now add the bookmark
        .Add Range:=oWA.Selection.Range, Name:=BM_Name
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
        
    setWordBookMark = True
End Function
'Private Sub enumerateWordBookMarks(oWA As Word.Application) 'early binding
Private Sub enumerateWordBookMarks(oWA As Object) 'late binding
'Dim BkMk As Word.Bookmark 'early binding
Dim BkMk As Object 'late binding

    For Each BkMk In oWA.ActiveDocument.Bookmarks
        Debug.Print BkMk.Name
    Next BkMk
End Sub
'Source: Adapted from http://www.vbaexpress.com/kb/getarticle.php?kb_id=562
'--------------------------------------------------------------------------
'Private Function BookmarkExists(oWD As Word.Document, sBookmark As String) As Boolean 'early binding
Private Function BookmarkExists(oWD As Object, sBookmark As String) As Boolean 'late binding

'Checks if a bookmark exists in the active document
     
    If oWD.Bookmarks.Exists(sBookmark) Then
        BookmarkExists = True
    Else
        BookmarkExists = False
    End If
End Function
'Private Function ISFormfield(oBMK As Word.Bookmark) As Boolean 'early binding
Private Function ISFormfield(oBMK As Object) As Boolean 'late binding
'Dim oFormField  As Word.FormField 'early binding
Dim oFormField As Object 'late binding
'Dim oWD As Word.Document 'early binding
Dim oWD As Object 'late binding
 
'Checks if bookmark IS a formfield
     
    Set oWD = oBMK.Parent
    
    If oWD.FormFields.Count = 0 Then
        ISFormfield = False
    Else
        For Each oFormField In oWD.FormFields()
            If oFormField.Name = oBMK.Name Then
                ISFormfield = True
            End If
        Next
    End If
End Function
'--------------------------------------------------------------------------
