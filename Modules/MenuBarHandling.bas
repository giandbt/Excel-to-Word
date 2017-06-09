Attribute VB_Name = "MenuBarHandling"
Option Explicit
Public Sub CreateMenu()
Dim myMenu As Variant
Dim MyBar As CommandBar
Dim MyPopup As CommandBarPopup
Dim MyButton As CommandBarButton

'Commandbar stuff - floating bar stuff

    For Each myMenu In Application.CommandBars 'enumerate controls to see if BackTrack is already loaded
        If myMenu.Name = "ExcelToWord!" Then 'MenuBar already exists - delete it and then let the app fall thru to recreate "new" version
            myMenu.Delete
        Else
            'do nothing
        End If
    Next myMenu
        
    Set MyBar = CommandBars.Add(Name:="ExcelToWord!", Position:=msoBarFloating, temporary:=True)
    
    With MyBar
        .Top = 175
        .Left = 850

        Set MyPopup = .Controls.Add(Type:=msoControlPopup)
        With MyPopup
            .Caption = "ExcelToWord!" 'change to suit
            .BeginGroup = True
            Set MyButton = .Controls.Add(Type:=msoControlButton)
            With MyButton
                .Caption = "&Configuration Options" 'enter name of macro to run
                .Style = msoButtonCaption
                ''' msoButtonAutomatic, msoButtonIcon, msoButtonCaption, or msoButtonIconandCaption
                .BeginGroup = True
                .OnAction = "showConfigurator" 'macro to be run
            End With
            Set MyButton = .Controls.Add(Type:=msoControlButton)
            With MyButton
                .Caption = "&Generate Word Bookmarks" 'enter name of macro to run
                .Style = msoButtonCaption
                ''' msoButtonAutomatic, msoButtonIcon, msoButtonCaption, or msoButtonIconandCaption
                .BeginGroup = False
                .OnAction = "generateWordBookmarks" 'macro to be run
            End With
            Set MyButton = .Controls.Add(Type:=msoControlButton)
            With MyButton
                .Caption = "&Update Word with Excel Data" 'enter name of macro to run
                .Style = msoButtonCaption
                ''' msoButtonAutomatic, msoButtonIcon, msoButtonCaption, or msoButtonIconandCaption
                .BeginGroup = False
                .OnAction = "updateWordFromExcel" 'macro to be run
            End With
            Set MyButton = .Controls.Add(Type:=msoControlButton)
            With MyButton
                .Caption = "&Name Embedded Object" 'enter name of macro to run
                .Style = msoButtonCaption
                ''' msoButtonAutomatic, msoButtonIcon, msoButtonCaption, or msoButtonIconandCaption
                .BeginGroup = True
                .OnAction = "nameEmbeddedObject" 'macro to be run
            End With
            Set MyButton = .Controls.Add(Type:=msoControlButton)
            With MyButton
                .Caption = "&Exit"
                .Style = msoButtonCaption
                .BeginGroup = True
                .OnAction = "ExcelToWord_UserTerminate"
            End With
            'this code to add another button
'            Set MyButton = .Controls.Add(Type:=msoControlButton)
'            With MyButton
'                .Caption = "&XXXXXXX"
'                .Style = msoButtonCaption
'                .BeginGroup = False
'                .OnAction = "XXXXXXX"
'            End With
        End With
   
    .Width = 150
    .Visible = True

    End With
End Sub
