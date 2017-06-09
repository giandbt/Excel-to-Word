VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Workbook_Open()

    If appVer < 12 Then 'check application version.  < 12 means earlier than Excel 2007, so use commandbar approach
        Call CreateMenu
    End If

End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If unInstalling Then Cancel = True
    
    If Not userTerminate Then
        Call ExcelToWord_Terminate
    End If
End Sub
Private Sub Workbook_AddinUninstall()
    userTerminate = True
End Sub
