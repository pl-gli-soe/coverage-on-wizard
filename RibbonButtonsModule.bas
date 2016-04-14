Attribute VB_Name = "RibbonButtonsModule"
Public Sub run_cow(ictrl As IRibbonControl)
    
    
    Dim ich As InitConfigHandler
    Set ich = New InitConfigHandler
    With CfgUserForm
    
    
        ' zanim show
        ' ====================================
        .CheckBoxPUSes.Value = False
        .CheckBoxRECV.Value = False
        .CheckBoxRQMs.Value = False
        
        .OptionButtonMGO.Value = False
        .OptionButtonPUS_MGO.Value = False
        .OptionButtonMIXED.Value = False
        
        .OptionButtonMGO.Value = False
        .OptionButtonWGEN.Value = False
        .OptionButtonCbalFromWizard.Value = False
        
        
        .CheckBoxRunCov.Value = False
        .CheckBoxRunFlats.Value = False
        
        ' cov type
        ' ====================================
        .OptionButtonPUSesFromMGO.Value = False
        .OptionButtonPUSesFromWiz.Value = False
        
        
        ' ====================================
        .CheckBoxCoordList.Value = False
        
        ' ====================================
        .connectWithExternalIch ich
        .Show
        
    End With
    
    
    Set ich = Nothing
    
    
    MsgBox "ready!"
End Sub


