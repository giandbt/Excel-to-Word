Attribute VB_Name = "ConfiguratorHandling"
Option Explicit
'Configurator Options --------------------
Public Const CONFIG_SETTINGS = "ETW_strWD_TemplOpt,ETW_bWD_Table,ETW_strWD_TemplFile,ETW_strWD_TemplateBMFile,ETW_strXL_TemplOpt,ETW_strXL_TemplOptShapePaste,ETW_strXL_TemplOptCell,ETW_bXL_SpanWorkbook,ETW_bXL_Increment,ETW_strXL_RefCounter,ETW_strXL_RefStart,ETW_strXL_RefEnd,ETW_bAftUpdPrint,ETW_bAftUpdPDF,ETW_bAftUpdSave,ETW_strAftUpdEmail,ETW_bAftUpdDelete,ETW_bAftUpdPreview,ETW_strWD_DocPath,ETW_strWD_DocFile,ETW_bSaveConfig"
Public Const RANGE_REF = "ETW_strXL_RefCounter,ETW_strXL_RefStart,ETW_strXL_RefEnd"
Public Const ChartShapeImagePasteOptions = "wdPasteBitmap,wdPasteDeviceIndependentBitmap,wdPasteEnhancedMetafile,wdPasteMetafilePicture,wdPasteOLEObject"
Public Const EmailWordOrPDFOptions = ",eWord,ePDF"
Public iXL_TemplOptShapePaste As Integer
Public Const adjacent = "Left,Above,Right,Below" '0,1,2,3.  0-Default = Left
Public Const CONFIG_SCOPE = "ETW_ConfiguratorScope"
Public Const CONFIG_SHEET = "ETW_ConfigSheet"
Public Const NAME_VISIBLE = False 'whether config names are visible or not - TRUE for debug purposes, only
Public strWD_TemplOpt As String 'Word template options:  User created "OWN", "GENERIC", or "INTELLIGENT" bookmarks
Public bWD_Table As Boolean 'TRUE: User has a 1-to-many row table, with bookmarks indicators embedded.  Option to delete empty rows from table, during processing
Public strWD_TemplFile As String 'Original word template - the starting point
Public strWD_TemplateBMFile As String 'Generated Word template, with bookmarks either user/system generated
Public strXL_TemplOpt As String 'Excel template options: User created named ranges/shapes, bookmark indicators to the left/above/right/below of data, or both
Public strXL_TemplOptShapePaste As String 'option for Picture or OLE Object link to Excel with shape paste (e.g., chart)
Public strXL_TemplOptCell As String 'bookmark indicators to the left, above, right, or below the data
Public bXL_SpanWorkbook As Boolean 'True - configuration options span entire workbook, as opposed to active sheet
Public bXL_Increment As Boolean 'True - Update Word from Excel will cycle based on counter, start & end point
Public strXL_RefCounter As String
Public strXL_RefStart As String
Public strXL_RefEnd As String
Public bAftUpdPrint As Boolean 'True - print after Word document is updated
Public bAftUpdPDF As Boolean 'True - PDF file will be extracted after Word template is updated
Public bAftUpdSave As Boolean 'True - Word document will be saved from Word template
Public bAftUpdDelete As Boolean 'True - Word document will be deleted after the update process (e.g., after printing or PDF process, etc.
Public strAftUpdEmail As String 'ePDF or eWord - Will email the PDF or Word output, as selected
Public bAftUpdPreview As Boolean 'True - just preview Word Draft after generation (only the 1st generated document)
Public strWD_DocPath As String 'path for saving Word document
Public strWD_DocFile As String 'Word document filename
Public bSaveConfig As Boolean 'True - save Configuration Options for next step
'-----------------------------------------
Public Const WORDDOC_PATH = "ETW_WordDocPath" 'stores the last path the user selected when browsing to set a word document path
Public Const WORDTMPL_PATH = "ETW_WordTemplPath" 'stores the last path the user selected when browsing for a word template
Public Const oWA_VISIBLE = False 'True - Word application will be visible during automation
Public Const SPACE_CHAR = " "
Public Const QUOTE_CHAR = "'"
Public closeOut As Boolean
Public Const VERSION_NO = "v1.1"
'-------------------------------- Late Binding variables needed ------------------------------
Public Const wdPasteEnhancedMetafile = 9
Public Const wdPasteBitmap = 4
Public Const wdPasteDeviceIndependentBitmap = 5
Public Const wdPasteMetaFilePicture = 3
Public Const wdPasteOLEObject = 0
Public Const wdGoToBookmark = -1
Public Const wdInLine = 0
Public Const wdExportFormatPDF = 17
Public Const wdExportOptimizeForPrint = 0
Public Const wdExportAllDocument = 0
Public Const wdExportDocumentContent = 0
Public Const wdExportCreateNoBookmarks = 0
Public Const wdRelativeHorizontalPositionColumn = 2
Public Const wdRelativeVerticalPositionParagraph = 2
Public Const wdRelativeHorizontalSizePage = 1
Public Const wdRelativeVerticalSizePage = 1
Public Const wdShapePositionRelativeNone = -999999
Public Const wdShapeSizeRelativeNone = -999999
Public Const wdWrapBoth = 0
Public Const wdFindContinue = 1
Public Const wdReplaceOne = 1
Public Const wdSortByName = 0
Public Sub showConfigurator(Optional control As Object) 'IRibbonControl

    If Application.Workbooks.Count = 0 Then
        MsgBox "No files open to process"
        Exit Sub
    End If
    
    If ActiveSheet.Type <> xlWorksheet Then
        MsgBox "You can only run ExcelToWord! functions from Excel Worksheets (e.g., Not from Chart Sheets, etc.)", vbCritical
    Else
        Load Configurator
        Configurator.Show
    End If
End Sub
Public Sub initializeConfiguratorOptions()
Dim strNamedScope As String
Dim objWkbSht As Object
Dim tmpVar As Variant
Dim bEvalSheet As Boolean

    'Initial Configuration Settings into Public Variables
    On Error Resume Next
    strNamedScope = myEvaluate(CONFIG_SCOPE)
    bEvalSheet = myEvaluate(CONFIG_SHEET)
    On Error GoTo 0
    
    'determine if any Configuration Settings exist
    If (strNamedScope = "Worksheet" And bEvalSheet) Or (strNamedScope = "Workbook" And Not bEvalSheet) Then 'there are settings at Workbook or this sheet's scope
        Call setPublicVariables
    Else 'there were no saved settings in the Workbook, or on the Active Sheet
        Call baseInitialization
    End If
End Sub
Public Sub setPublicVariables()
Dim varConfig As Variant
Dim i As Integer
Dim tmpVar As Variant

    varConfig = Split(CONFIG_SETTINGS, ",")
    For i = 0 To UBound(varConfig)
        On Error Resume Next
        If InStr(UCase(RANGE_REF), UCase(varConfig(i))) <> 0 Then 'get range object, as opposed to string value
            Set tmpVar = myEvaluate(varConfig(i))
            If Err.Number <> 0 Then
                GoTo errHandler
            End If
            tmpVar = "'" & tmpVar.Worksheet.Name & "'!" & tmpVar.Address
        Else
            tmpVar = myEvaluate(varConfig(i))
        End If
    
        Call setVar(varConfig(i), tmpVar)
errHandler:
        On Error GoTo 0
    Next i
End Sub
Private Sub baseInitialization()
    strWD_TemplOpt = "OWN"
    bWD_Table = False
    strWD_TemplFile = vbNullString
    strWD_TemplateBMFile = vbNullString
    strXL_TemplOpt = "RANGE"
    strXL_TemplOptShapePaste = "wdPasteEnhancedMetafile"
    iXL_TemplOptShapePaste = wdPasteEnhancedMetafile
    strXL_TemplOptCell = "Left"
    bXL_SpanWorkbook = IIf(myEvaluate(CONFIG_SCOPE) = "Worksheet", False, True) 'in case another sheet already has options saved
    bXL_Increment = False
    strXL_RefCounter = vbNullString
    strXL_RefStart = vbNullString
    strXL_RefEnd = vbNullString
    bAftUpdPrint = False
    bAftUpdPDF = False
    bAftUpdSave = True
    strAftUpdEmail = vbNullString
    bAftUpdPreview = False
    bAftUpdDelete = False
    strWD_DocPath = vbNullString
    strWD_DocFile = vbNullString
    bSaveConfig = False
End Sub
Public Function validateFileFolderSelection(ByVal fName As String, fType As String, src As String, bFolderOnly As Boolean) As Boolean
'Dim FSO As FileSystemObject 'early binding
Dim FSO As Object 'late binding
   

    'Set FSO = New FileSystemObject 'early binding
    Set FSO = CreateObject("Scripting.FileSystemObject") 'late binding

    validateFileFolderSelection = True
    
    'Test for word or excel filename & that the file exists
    If Trim(fName) = vbNullString Then
        validateFileFolderSelection = False
    ElseIf bFolderOnly Then
        If Not FSO.FolderExists(fName) Then
            validateFileFolderSelection = False
        End If
    ElseIf Not FSO.fileExists(fName) Then
            validateFileFolderSelection = False
    End If
    
End Function
Public Function browseForTemplate(strPath As String, strFilter1 As String, strFilter2, strTitle As String, bgetFolderOnly) As String
Dim dialogFile As FileDialog
Dim fName As String

    ' Open the file dialog
    Set dialogFile = Application.FileDialog(IIf(bgetFolderOnly, msoFileDialogFolderPicker, msoFileDialogFilePicker))
    With dialogFile
        If Not bgetFolderOnly Then
            .Filters.Clear
            .Filters.Add strFilter1, strFilter2, 1
        End If
        
        .AllowMultiSelect = False
        .InitialView = msoFileDialogViewDetails
        .InitialFileName = strPath
        .Title = strTitle
        .Show
    End With
    If dialogFile.SelectedItems.Count > 0 Then
        browseForTemplate = dialogFile.SelectedItems(1)
    Else
        browseForTemplate = ""
    End If
    
    'cleanup
    Set dialogFile = Nothing
End Function
Public Function getVar(strRef As Variant) As Variant

    Select Case strRef
        Case "ETW_strWD_TemplOpt":              getVar = strWD_TemplOpt
        Case "ETW_bWD_Table":                   getVar = bWD_Table
        Case "ETW_strWD_TemplFile":             getVar = strWD_TemplFile
        Case "ETW_strWD_TemplateBMFile":        getVar = strWD_TemplateBMFile
        Case "ETW_strXL_TemplOpt":              getVar = strXL_TemplOpt
        Case "ETW_strXL_TemplOptShapePaste":    getVar = strXL_TemplOptShapePaste
        Case "ETW_strXL_TemplOptCell":          getVar = strXL_TemplOptCell
        Case "ETW_bXL_SpanWorkbook":            getVar = bXL_SpanWorkbook
        Case "ETW_bXL_Increment":               getVar = bXL_Increment
        Case "ETW_strXL_RefCounter":            getVar = strXL_RefCounter
        Case "ETW_strXL_RefStart":              getVar = strXL_RefStart
        Case "ETW_strXL_RefEnd":                getVar = strXL_RefEnd
        Case "ETW_bAftUpdPrint":                getVar = bAftUpdPrint
        Case "ETW_bAftUpdPDF":                  getVar = bAftUpdPDF
        Case "ETW_bAftUpdSave":                 getVar = bAftUpdSave
        Case "ETW_bAftUpdDelete":               getVar = bAftUpdDelete
        Case "ETW_strAftUpdEmail":              getVar = strAftUpdEmail
        Case "ETW_bAftUpdPreview":              getVar = bAftUpdPreview
        Case "ETW_strWD_DocPath":               getVar = strWD_DocPath
        Case "ETW_strWD_DocFile":               getVar = strWD_DocFile
        Case "ETW_bSaveConfig":                 getVar = bSaveConfig
    End Select
    
End Function
Private Function setVar(strRef As Variant, myVal As Variant) As String

    On Error Resume Next
    
    Select Case strRef
        Case "ETW_strWD_TemplOpt":              strWD_TemplOpt = myVal
        Case "ETW_bWD_Table":                   bWD_Table = myVal
        Case "ETW_strWD_TemplFile":             strWD_TemplFile = myVal
        Case "ETW_strWD_TemplateBMFile":        strWD_TemplateBMFile = myVal
        Case "ETW_strXL_TemplOpt":              strXL_TemplOpt = myVal
        Case "ETW_strXL_TemplOptShapePaste":    strXL_TemplOptShapePaste = myVal
                                                iXL_TemplOptShapePaste = setVarShapePaste(myVal)
        Case "ETW_strXL_TemplOptCell":          strXL_TemplOptCell = myVal
        Case "ETW_bXL_SpanWorkbook":            bXL_SpanWorkbook = myVal
        Case "ETW_bXL_Increment":               bXL_Increment = myVal
        Case "ETW_strXL_RefCounter":            strXL_RefCounter = myVal
        Case "ETW_strXL_RefStart":              strXL_RefStart = myVal
        Case "ETW_strXL_RefEnd":                strXL_RefEnd = myVal
        Case "ETW_bAftUpdPrint":                bAftUpdPrint = myVal
        Case "ETW_bAftUpdPDF":                  bAftUpdPDF = myVal
        Case "ETW_bAftUpdSave":                 bAftUpdSave = myVal
        Case "ETW_bAftUpdDelete":               bAftUpdDelete = myVal
        Case "ETW_strAftUpdEmail":              strAftUpdEmail = myVal
        Case "ETW_bAftUpdPreview":              bAftUpdPreview = myVal
        Case "ETW_strWD_DocPath":               strWD_DocPath = myVal
        Case "ETW_strWD_DocFile":               strWD_DocFile = myVal
        Case "ETW_bSaveConfig":                 bSaveConfig = myVal
    End Select
    
    On Error GoTo 0
    
End Function
Private Function setVarShapePaste(strOpt As Variant) As Integer
'I selected what I thought the most relevant of paste options in Word parlance, with several physical picture options, and a link option to the original workbook

    Select Case strOpt
        Case "wdPasteBitmap":                   setVarShapePaste = wdPasteBitmap
        Case "wdPasteDeviceIndependentBitmap":  setVarShapePaste = wdPasteDeviceIndependentBitmap
        Case "wdPasteEnhancedMetafile":         setVarShapePaste = wdPasteEnhancedMetafile
        Case "wdPasteMetafilePicture":          setVarShapePaste = wdPasteMetaFilePicture
        Case "wdPasteOLEObject":                setVarShapePaste = wdPasteOLEObject
    End Select
End Function
Public Sub resetConfigurator()
Dim varConfig As Variant
Dim wks As Worksheet
Dim i As Integer

    varConfig = Split(CONFIG_SETTINGS, ",")
    
    'deletes all Configurator references in Workbook - at Workbook and Sheet-level
    On Error Resume Next
    ActiveWorkbook.Names(CONFIG_SCOPE).Delete
    ActiveWorkbook.Names(WORDDOC_PATH).Delete
    ActiveWorkbook.Names(WORDTMPL_PATH).Delete
    
    'delete at Workbook Scope, if any
    For i = 0 To UBound(varConfig)
        ActiveWorkbook.Names(varConfig(i)).Delete
    Next i
    
    'delete at Worksheet Scope, if any
    For Each wks In ActiveWorkbook.Worksheets
        wks.Names(CONFIG_SHEET).Delete
        For i = 0 To UBound(varConfig)
            wks.Names(varConfig(i)).Delete
        Next i
    Next wks
    
    On Error GoTo 0
End Sub

