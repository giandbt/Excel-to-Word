Attribute VB_Name = "AddinsInstall"
Option Explicit
'Source: Significant adaptation from original by rmclean Feb 2006 @ http://www.blog.methodsinexcel.co.uk/2006/03/02/click-to-install-addin/
'Purpose : Installs Addin to local addins folder or user selected folder
Sub auto_open()
Dim gsFilename As String
Dim gsAppName As String
Dim AddinInstallPath As String
Dim bInstalled As Boolean
Dim chkWkb As Workbook
Dim xMsg As Long
    
    If ThisWorkbook.FileFormat <> xlOpenXMLAddIn And _
        ThisWorkbook.FileFormat <> xlAddIn Then 'not an add-in file, so abort
            Exit Sub
    End If
    
    gsAppName = "ExcelToWord!" 'Whatever name you want to show up in the list of Excel Add-In Options for this Add-In
    gsFilename = gsAppName & Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStr(ThisWorkbook.Name, ".") + 1)
    
    If ThisWorkbook.Name = gsFilename Then Exit Sub 'just installed.  Install files will have version numbers to the right of ExcelToWord! text
    
    On Error Resume Next
    AddinInstallPath = GetSetting(appname:=gsAppName, section:="User Addin", Key:="InstallPath")
    On Error GoTo 0
    
    If AddinInstallPath = "" Then 'presume to re-install at same location as last save
        xMsg = MsgBox("Would you like to install this add-in, or just run it one time?", vbYesNo, "Hit YES to install, NO to just run")
        If xMsg = vbYes Then
            AddinInstallPath = Application.UserLibraryPath
            xMsg = MsgBox("Install at default user library path: " & AddinInstallPath & "?", vbYesNo, "Hit YES to proceed, NO to prompt for directory")
            If xMsg = vbNo Then
                'get path for add-in installation
                With Application.FileDialog(msoFileDialogFolderPicker)
                    .AllowMultiSelect = False
                    .InitialView = msoFileDialogViewLargeIcons
                    .InitialFileName = Application.UserLibraryPath
                    If .Show <> -1 Then
                        MsgBox "No folder selected! Aborting Install - Instead will run Add-in one time..."
                        GoTo gracefulExit 'just let the AddIn proceed without installation
                    Else
                        AddinInstallPath = .SelectedItems(1) & "\"
                    End If
                End With
            End If
        Else
            GoTo gracefulExit 'just let the AddIn proceed without installation
        End If
    End If

    On Error GoTo ErrorHander
  
    If UCase(ThisWorkbook.FullName) = UCase(AddinInstallPath & gsFilename) Then
        On Error Resume Next
        AddIns.Add (AddinInstallPath & gsFilename)
        AddIns(gsAppName).Installed = True
        bInstalled = True
        SaveSetting appname:=gsAppName, section:="User Addin", Key:="Installed", Setting:="True"
        SaveSetting appname:=gsAppName, section:="User Addin", Key:="InstallPath", Setting:=AddinInstallPath
        GoTo gracefulExit
    Else
        'ensure at least one workbook open
        If Application.Workbooks.Count = 0 Then Application.Workbooks.Add
        
        'close any workbook with the same name
        On Error Resume Next
        Set chkWkb = Application.Workbooks(gsFilename)
        If Not chkWkb Is Nothing Then
            chkWkb.Close SaveChanges:=False
        End If
        
        On Error GoTo ErrorHander
        
        'take care of existing add-in with same name at same save location as new installation
        If Dir(AddinInstallPath & gsFilename) <> "" Then
            xMsg = MsgBox("You want to replace existing add-in file in the " & vbNewLine & AddinInstallPath & _
                vbNewLine & " directory?", vbYesNo, "Hit YES to replace, NO to abort")
            If xMsg = vbNo Then
                GoTo gracefulExit
            Else
                Application.DisplayAlerts = False
                Application.ScreenUpdating = False
                
                On Error Resume Next
                Application.AddIns(gsAppName).Installed = False
                Kill AddinInstallPath & gsFilename
                SaveSetting appname:=gsAppName, section:="User Addin", Key:="Installed", Setting:="False"
            End If
        End If
        
        ThisWorkbook.IsAddin = True
        If ThisWorkbook.FileFormat = xlOpenXMLAddIn Then 'save as Excel 2007/2010 .XLAM file
            ThisWorkbook.SaveAs Filename:=AddinInstallPath & gsFilename, FileFormat:=xlOpenXMLAddIn
        Else 'save as Excel 2003 .XLA file
            ThisWorkbook.SaveAs Filename:=AddinInstallPath & gsFilename, FileFormat:=xlAddIn
        End If
        
        AddIns.Add (AddinInstallPath & gsFilename)
        AddIns(gsAppName).Installed = True
        bInstalled = True
        SaveSetting appname:=gsAppName, section:="User Addin", Key:="Installed", Setting:="True"
        SaveSetting appname:=gsAppName, section:="User Addin", Key:="InstallPath", Setting:=AddinInstallPath
    End If
    
    GoTo gracefulExit
    
ErrorHander:
    MsgBox "An error happened, when i wrote this code I didn’t think this would happen. The error was: " _
    & vbNewLine & vbNewLine _
    & Err.Description & vbNewLine & vbNewLine _
    & "Can you tell the author about this. Thanks", Title:=gsAppName, Buttons:=vbOKOnly
    
gracefulExit:
    If bInstalled Then ThisWorkbook.Close SaveChanges:=False
End Sub
Sub addInUninstall()
Dim gsFilename As String
Dim gsAppName As String
Dim AddinInstallPath As String
Dim chkWkb As Workbook
Dim xMsg As Long
    
    gsAppName = "ExcelToWord!" 'Whatever name you want to show up in the list of Excel Add-In Options for this Add-In
    gsFilename = gsAppName & Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStr(ThisWorkbook.Name, ".") + 1)
    
    On Error Resume Next
    
    AddinInstallPath = GetSetting(appname:=gsAppName, section:="User Addin", Key:="InstallPath")
    If AddinInstallPath = "" Then AddinInstallPath = Application.UserLibraryPath
    
    ThisWorkbook.IsAddin = False
    DeleteSetting appname:=gsAppName
    
    unInstalling = True
    Application.AddIns(gsAppName).Installed = False
    unInstalling = False
    
    ThisWorkbook.Close SaveChanges:=False
End Sub
