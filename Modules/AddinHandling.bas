Attribute VB_Name = "AddinHandling"
Option Explicit
'in a public module, the following declarations (so you can close the menu bar later...
Public userTerminate As Boolean
Public unInstalling As Boolean
Public Sub ExcelToWord_Terminate()
Dim myAddin As AddIn

    On Error Resume Next 'if exit commandbar menu, then this fires, then workbook_BeforeClose fires, so would get an error the second time around
    
    Application.CommandBars("ExcelToWord!").Delete

    ThisWorkbook.Close SaveChanges:=False
    
    On Error GoTo 0
End Sub
Public Sub ExcelToWord_UserTerminate(Optional control As Object) 'IRibbonControl
    userTerminate = True
    MsgBox "Shutting Down ExcelToWord! Version: " & VERSION_NO
    Call ExcelToWord_Terminate
End Sub
Public Sub tempUnmakeAddin()
     ThisWorkbook.IsAddin = False
End Sub

