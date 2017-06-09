VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Configurator 
   Caption         =   "ExcelToWord! Configuration Settings"
   ClientHeight    =   11355
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7860
   OleObjectBlob   =   "Configurator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Configurator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lastFocus As Object

Private Sub cbEmail_Change()
    If cbEmail = "" Then
        lblMessage2.Visible = False
    Else
        lblMessage2.Visible = True
    End If
End Sub

Private Sub cbSave_Click()
    Me.cbSave = True
End Sub

Private Sub cmdBrowseWordDocument_Click()
Dim tmpResult As String
Dim strPath As String

    If Me.tbWordDocumentPath <> "" Then
        strPath = Me.tbWordDocumentPath & "\"
    Else
        strPath = myEvaluate(WORDDOC_PATH)
        If strPath = "" Then strPath = ActiveWorkbook.path & "\"
    End If
    
    tmpResult = browseForTemplate(strPath, "Word files", "*.doc; *.docx; *.docm", "Select path for generating Word document", True)
    If tmpResult <> "" Then
        Me.tbWordDocumentPath = tmpResult
        Application.Names.Add Name:=WORDDOC_PATH, RefersTo:=tmpResult, Visible:=NAME_VISIBLE
    End If
End Sub

Private Sub cmdBrowseWordTemplate_Click()
Dim tmpResult As String
Dim strPath As String

    If Me.tbWordTemplate <> "" Then
        strPath = getPathFromPathFName(Me.tbWordTemplate)
    Else
        strPath = myEvaluate(WORDTMPL_PATH)
        If strPath = "" Then strPath = Application.TemplatesPath
    End If
       
    tmpResult = browseForTemplate(strPath, "Word template files", "*.dot; *.dotx; *.dotm", "Select Word template", False)
    If tmpResult <> "" Then
        Me.tbWordTemplate = tmpResult
        Application.Names.Add Name:=WORDTMPL_PATH, RefersTo:=tmpResult, Visible:=NAME_VISIBLE
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdReset_Click()
Dim xMsg As Long

    xMsg = MsgBox("Reset Configuration Settings for Entire Workbook?", vbYesNo, "Hit YES to Reset Everything")
    If xMsg = vbYes Then
        Call resetConfigurator
        Unload Me
    End If
End Sub


Private Sub obRangeNames_Click()
    Me.cbAdjacent.Visible = False
    Me.cbShapePaste.Visible = True
End Sub
Private Sub obBookmarkIndicators_Click()
    Me.cbAdjacent.Visible = True
    Me.cbShapePaste.Visible = False
End Sub
Private Sub obBookmarkAndRangeNames_Click()
    Me.cbAdjacent.Visible = True
    Me.cbShapePaste.Visible = True
End Sub

Private Sub wordStuffVisible(bVisible As Boolean)
    Me.tbWordDocumentPath.Enabled = bVisible
    Me.tbWordDocumentName.Enabled = bVisible
    Me.tbWordDocumentPath.Visible = bVisible
    Me.tbWordDocumentName.Visible = bVisible
    Me.lblWordPath.Visible = bVisible
    Me.lblWordFile.Visible = bVisible
    Me.cmdBrowseWordDocument.Enabled = bVisible
End Sub

Private Sub obPreviewAfterUpdate_Click()
    If Me.cbPDFAfterUpdate Then
        Call wordStuffVisible(True)
    Else
        Call wordStuffVisible(False)
    End If
End Sub

Private Sub cbPDFAfterUpdate_Click()
    If Me.cbPDFAfterUpdate Then
        If Not Me.obSaveAfterUpdate Then
            Me.obSaveAfterUpdate = True
        End If
    End If
End Sub

Private Sub obDeleteAfterUpdate_Click()
    If Not Me.cbPDFAfterUpdate Then
        Call wordStuffVisible(False)
    Else
        Call wordStuffVisible(True)
    End If
    
End Sub

Private Sub obSaveAfterUpdate_Click()
    Call wordStuffVisible(True)
End Sub

Private Sub cbIncrement_Click()
    Call showIncrement
End Sub
Private Sub showIncrement()
Dim bVisible As Boolean

    bVisible = Me.cbIncrement
    
    Me.refCounter.Visible = bVisible
    Me.refStart.Visible = bVisible
    Me.refEnd.Visible = bVisible
    Me.lblFrom.Visible = bVisible
    Me.lblTo.Visible = bVisible
    Me.lblStartCtr.Visible = bVisible
    Me.lblEndCtr.Visible = bVisible
    Me.lblMessage.Visible = bVisible
    
End Sub

Private Sub refCounter_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim chkRef As Range

    'must be a range reference
    If Not closeOut And Me.refCounter <> "" Then 'if Userform has been cancelled, then skip this check
    
        On Error Resume Next 'test range selected for increment counter
        Set chkRef = Range(Me.refCounter)
        If Err.Number <> 0 Then
            MsgBox "You must select a valid range for your increment counter", vbCritical
            Cancel = True
        ElseIf chkRef.Count > 1 Then
            MsgBox "You must only select 1 cell reference for the increment counter", vbCritical
            Cancel = True
        End If
        On Error GoTo 0
    End If
    
End Sub

Private Sub refStart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim chkRef As Range

    'must be a range reference
    If Not closeOut And Me.refStart <> "" Then 'if Userform has been cancelled, then skip this check
    
        On Error Resume Next 'test range selected for increment counter
        Set chkRef = Range(Me.refStart)
        If Err.Number <> 0 Then
            MsgBox "You must select a valid range for your increment counter", vbCritical
            Cancel = True
        ElseIf chkRef.Count > 1 Then
            MsgBox "You must only select 1 cell reference for the increment counter", vbCritical
            Cancel = True
        Else
            Me.lblStartCtr.Caption = chkRef.Value
        End If
        On Error GoTo 0
    End If
    
End Sub

Private Sub refEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim chkRef As Range

    'must be integer or range reference
    If Not closeOut And Me.refEnd <> "" Then 'if Userform has been cancelled, then skip this check
    
        On Error Resume Next 'test range selected for increment counter
        Set chkRef = Range(Me.refEnd.Value)
        If Err.Number <> 0 Then
            MsgBox "You must select a valid range for your increment counter", vbCritical
            Cancel = True
        ElseIf chkRef.Count > 1 Then
            MsgBox "You must only select 1 cell reference for the increment counter", vbCritical
            Cancel = True
        Else
            Me.lblEndCtr.Caption = chkRef.Value
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub showIncrementCtrs()

    On Error Resume Next
    Me.lblStartCtr.Caption = Range(Me.refStart).Value
    Me.lblEndCtr.Caption = Range(Me.refEnd).Value
    On Error GoTo 0
    
End Sub
   
Private Sub tbWordTemplate_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not closeOut And Me.tbWordTemplate <> "" Then  'if Userform has been cancelled, then skip this check
        If Not validateFileFolderSelection(Me.tbWordTemplate, "Word", "template", False) Then
            MsgBox "The path\filename entered does not exist" & Chr(10) & Chr(10) & Me.tbWordTemplate & Chr(10) & Chr(10) & _
                "Fix entry, or delete entry and BROWSE for file", vbOKOnly, "Configurator Error"
                
            Cancel = True
        End If
    End If
End Sub

Private Sub tbWordDocumentPath_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not closeOut And Me.tbWordDocumentPath <> "" Then 'if Userform has been cancelled, then skip this check
        
        If Right(Me.tbWordDocumentPath, 1) = "\" Then Me.tbWordDocumentPath = Left(Me.tbWordDocumentPath, Len(Me.tbWordDocumentPath) - 1)
        
        If Not validateFileFolderSelection(Me.tbWordDocumentPath, "Word", "document", True) Then
            MsgBox "The path entered does not exist" & Chr(10) & Chr(10) & Me.tbWordDocumentPath & Chr(10) & Chr(10) & _
                "Fix entry, or delete entry and BROWSE for file", vbOKOnly, "Configurator Error"
                
            Cancel = True
        End If
    End If
    
End Sub

Private Sub tbWordDocumentName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim strPathFName As String

    strPathFName = Me.tbWordDocumentPath & "\" & Me.tbWordDocumentName
    If Not closeOut And Me.tbWordDocumentName <> "" Then 'if userform has been cancelled, then skip this check
    
        If Not validateFileFolderSelection(strPathFName, "Word", "document", True) Then 'check if the file exists at this path, already
            If Not IsLegalFileName(Me.tbWordDocumentName) Then 'check construction of path and filename - any illegal characters?
                MsgBox "Invalid Word filename" & Chr(10) & Chr(10) & Me.tbWordDocumentName & Chr(10) & Chr(10) & _
                    "Fix entry", vbOKOnly, "Configurator Error"
                
                Cancel = True
            End If
        End If
        
    End If
    
End Sub
Private Sub Userform_Initialize()
Dim tmpVar As Variant
Dim i As Integer

    closeOut = False
    
    'load Public Variables with defaults or from named ranges saved
    Call initializeConfiguratorOptions
    
    'Now, initialize Configurator with Public variables
    Select Case strWD_TemplOpt:
        Case "OWN": Me.obMyOwnBM = True
        Case "GENERIC": Me.obGenericBM = True
        Case "INTELLIGENT": Me.obIntelBM = True
    End Select
    
    Me.cbTable = bWD_Table
    Me.tbWordTemplate = strWD_TemplFile

    Me.cbAdjacent.Visible = True
    Me.cbShapePaste.Visible = True
    
    Select Case strXL_TemplOpt:
        Case "RANGE":
            Me.obRangeNames = True
            Me.cbAdjacent.Visible = False
            Me.cbShapePaste.Visible = False
        Case "CELL":
            Me.obBookmarkIndicators = True
        Case "RANGE_AND_CELL":
            Me.obBookmarkAndRangeNames = True
    End Select
    
    tmpVar = Split(adjacent, ",")
    
    For i = 0 To UBound(tmpVar)
        Me.cbAdjacent.AddItem tmpVar(i), i
    Next i
    
    Me.cbAdjacent = strXL_TemplOptCell
    
    tmpVar = Split(ChartShapeImagePasteOptions, ",")
    
    For i = 0 To UBound(tmpVar)
        Me.cbShapePaste.AddItem tmpVar(i), i
    Next i

    Me.cbShapePaste = strXL_TemplOptShapePaste
    
    tmpVar = Split(EmailWordOrPDFOptions, ",")
    For i = 0 To UBound(tmpVar)
        Me.cbEmail.AddItem tmpVar(i), i
    Next i
    
    Me.cbEmail = strAftUpdEmail
    
    Me.cbExcelSpanWorkbook = bXL_SpanWorkbook
    Me.cbIncrement = bXL_Increment
    Me.refCounter = strXL_RefCounter
    Me.refStart = strXL_RefStart
    Me.refEnd = strXL_RefEnd
    Me.cbPrintAfterUpdate = bAftUpdPrint
    Me.cbPDFAfterUpdate = bAftUpdPDF
    Me.obSaveAfterUpdate = bAftUpdSave
    Me.obPreviewAfterUpdate = bAftUpdPreview
    Me.obDeleteAfterUpdate = Not (bAftUpdSave Or bAftUpdPreview)
    Me.tbWordDocumentPath = strWD_DocPath
    Me.tbWordDocumentName = strWD_DocFile
    Me.cbSave = bSaveConfig
    
    Call showIncrement
    Call showIncrementCtrs
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    closeOut = True
End Sub

Private Sub cmdSave_Click()
Dim validError As String
Dim bError As Boolean
Dim varConfig As Variant
Dim strNameScope As String
Dim objWkbSht As Object
Dim i As Integer
Dim xMsg As Long
Dim tmpVar As Variant
Dim strPathFName As String

    'Validation before Userform exit
    'Configurator settings saved at scope of ActiveSheet, to allow for different settings on different sheets, re: different Word templates/Excel data pulls might be possible
    
    'Test for scope change - if from Workbook to Sheet, or from Sheet to Workbook - validate then cleanup to desired state
    strNameScope = myEvaluate(CONFIG_SCOPE)
    
    If (strNameScope = "Workbook" And Not Me.cbExcelSpanWorkbook) Or _
        (strNameScope = "Worksheet" And Me.cbExcelSpanWorkbook) Then 'we have a misalignment which must be handled
    
        If Me.cbExcelSpanWorkbook Then
            xMsg = MsgBox("You have selected the scope for bookmark indicator searches to be the entire workbook, but you had previously selected " & _
                "Sheet-specific scope, and thus your configuration settings were saved at the Sheet level.  Would you like to reset the Workbook's " & _
                "configuration settings and go with Workbook level scoping?", vbYesNo, "Hit YES to Reset, then restart the setup, NO to return to the Configurator Userform")
        Else
            xMsg = MsgBox("You have selected the scope for bookmark indicator searches to be specific to this Worksheet, but you had previously " & _
                "selected entire Workbook-specific scope, and thus your configuration settings were saved at the Workbook level.  Would you like " & _
                "to reset the Workbook's configuration settings and go with Sheet level scoping?", vbYesNo, _
                "Hit YES to Reset, NO to return to the Configurator Userform")
        End If
        
        If xMsg = vbNo Then
            Exit Sub
        Else
            bXL_SpanWorkbook = Me.cbExcelSpanWorkbook
            strNameScope = IIf(bXL_SpanWorkbook, "Workbook", "Worksheet")
            Call resetConfigurator
            GoTo continueSaveProcess
        End If
        
    ElseIf strNameScope <> "" And Not Me.cbSave Then
        xMsg = MsgBox("You have previously saved configuration settings in this Workbook, and now you've selected that you don't want to save them.  " & _
            "Are you sure you want to proceed with resetting Configuration settings?", vbYesNo, "Hit YES to continue with reset, NO to return " & _
            "to the Configurator")
        If xMsg = vbNo Then
            Exit Sub
        End If
    End If
    
continueSaveProcess:

    'Test for word template filename & that the file exists
    If Trim(Me.tbWordTemplate) = vbNullString Then
        validError = "You must enter a valid Word Template path\filename"
    ElseIf Not validateFileFolderSelection(Me.tbWordTemplate, "Word", "template", False) Then
        validError = "The path\filename entered does not exist" & Chr(10) & Chr(10) & "[path\filename]: " & Me.tbWordTemplate
    End If
    If validError <> "" Then
        Set lastFocus = Me.tbWordTemplate
        GoTo backToUserform
    End If

    'Test for word document path & name
    If Me.obSaveAfterUpdate Or Me.cbPDFAfterUpdate Then
        If Trim(Me.tbWordDocumentPath) = vbNullString Or Trim(Me.tbWordDocumentPath) = vbNullString Then
            validError = "You must enter a valid Word document path"
        ElseIf Not validateFileFolderSelection(Me.tbWordDocumentPath, "Word/PDF", "document", True) Then
            validError = Me.tbWordDocumentPath & " The path\filename entered does not exist"
        End If
        If validError <> "" Then
            Set lastFocus = Me.tbWordDocumentPath
            GoTo backToUserform
        End If
        
        'Test for valid word document name
        strPathFName = Me.tbWordDocumentPath & "\" & Me.tbWordDocumentName
        If Me.tbWordDocumentName <> "" Then
            If Not validateFileFolderSelection(strPathFName, "Word/PDF", "document", True) Then 'check if the file exists at this path, already
                If Not IsLegalFileName(Me.tbWordDocumentName) Then 'check construction of path and filename - any illegal characters?
                    validError = "The Word/PDF Document name entered: " & Me.tbWordDocumentName & " is not a valid "
                End If
            End If
        Else
            validError = "You must enter a valid Word/PDF document filename"
        End If
        If validError <> "" Then
            Set lastFocus = Me.tbWordDocumentName
            GoTo backToUserform
        End If
    End If
    
    'Validate Incrementor variables
    If Me.cbIncrement And (Trim(Me.refCounter) = "" Or Trim(Me.refStart) = "" Or Trim(Me.refEnd) = "") Then
        validError = "You have selected the Incrementor Option, but have not identified the Start or End reference."
        If Trim(Me.refCounter) = "" Then
            Set lastFocus = Me.refCounter
        ElseIf Trim(Me.refStart) = "" Then
            Set lastFocus = Me.refStart
        Else
            Set lastFocus = Me.refEnd
        End If
        GoTo backToUserform
    End If
    
    'Validate eMail variable
    If Me.cbEmail = "ePDF" And (Not Me.cbPDFAfterUpdate) Then
        validError = "You selected to email the PDF, but did not select the option to Extract to PDF"
        GoTo backToUserform
    ElseIf Me.cbEmail = "eWord" And (Not obSaveAfterUpdate) Then
        validError = "You selected to email the Word document, but did not select the option to Save the Word document"
        GoTo backToUserform
    End If
    
    'initialize Public variables with Configuration Settings
    strWD_TemplOpt = IIf(Me.obMyOwnBM, "OWN", IIf(Me.obGenericBM, "GENERIC", "INTELLIGENT"))
    bWD_Table = Me.cbTable
    strWD_TemplFile = Me.tbWordTemplate
    strXL_TemplOpt = IIf(Me.obRangeNames, "RANGE", IIf(Me.obBookmarkIndicators, "CELL", "RANGE_AND_CELL"))
    strXL_TemplOptShapePaste = Me.cbShapePaste
    strXL_TemplOptCell = Me.cbAdjacent
    bXL_SpanWorkbook = Me.cbExcelSpanWorkbook
    bXL_Increment = Me.cbIncrement
    strXL_RefCounter = Me.refCounter
    strXL_RefStart = Me.refStart
    strXL_RefEnd = Me.refEnd
    bAftUpdPrint = Me.cbPrintAfterUpdate
    bAftUpdPDF = Me.cbPDFAfterUpdate
    bAftUpdSave = Me.obSaveAfterUpdate
    strAftUpdEmail = Me.cbEmail
    bAftUpdPreview = Me.obPreviewAfterUpdate
    bAftUpdDelete = Me.obDeleteAfterUpdate
    strWD_DocPath = Me.tbWordDocumentPath
    'strip any user-input extensions from the filename (as Word will automatically add based on the version of the document template)
    If InStr(Me.tbWordDocumentName, ".") <> 0 Then
    
        strWD_DocFile = Left(Me.tbWordDocumentName, InStr(Me.tbWordDocumentName, ".") - 1)
    Else
        strWD_DocFile = Me.tbWordDocumentName
    End If
    
    bSaveConfig = Me.cbSave
    
    varConfig = Split(CONFIG_SETTINGS, ",")
    Set objWkbSht = IIf(bXL_SpanWorkbook, ActiveWorkbook, ActiveSheet)
    strNameScope = IIf(bXL_SpanWorkbook, "Workbook", "Worksheet")
    
    If bSaveConfig Then
        'Create Range names for each Configurator variable - invisible
        On Error Resume Next 'to avoid error associated with Name having NULL string
        
        ActiveWorkbook.Names.Add Name:=CONFIG_SCOPE, RefersTo:=strNameScope, Visible:=NAME_VISIBLE
        If strNameScope = "Worksheet" Then
            objWkbSht.Names.Add Name:=CONFIG_SHEET, RefersTo:=True, Visible:=NAME_VISIBLE
        End If
        
        For i = 0 To UBound(varConfig)
            If InStr(UCase(RANGE_REF), UCase(varConfig(i))) <> 0 Then 'get range object, as opposed to string value
                Set tmpVar = Range(getVar(varConfig(i)))
            Else
                tmpVar = getVar(varConfig(i))
            End If
                
            objWkbSht.Names.Add Name:=varConfig(i), RefersTo:=tmpVar, Visible:=NAME_VISIBLE
        Next i
        
        On Error GoTo 0
    Else
        'Delete any range names that were created
        On Error Resume Next 'for names that didn't get created due to having NULL string
        
        ActiveWorkbook.Names(CONFIG_SCOPE).Delete
        For i = 0 To UBound(varConfig)
            objWkbSht.Names(varConfig(i)).Delete
        Next i
        
        On Error GoTo 0
    End If
    
    Unload Me
    
backToUserform:
    If validError <> "" Then MsgBox validError, vbOKOnly, "Configurator Error"
End Sub
