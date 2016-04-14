Attribute VB_Name = "RibbonButtonsModule"
Public Sub run_cow(ictrl As IRibbonControl)
    
    
    Dim ich As InitConfigHandler
    Set ich = New InitConfigHandler
    With CfgUserForm
    
    
        ' zanim show
        ' ====================================
        
        ' ' falty section ' '
        .CheckBoxPUSes.Value = False
        .CheckBoxCBALs.Value = False
        .CheckBoxRECV.Value = False
        .CheckBoxRQMs.Value = False
        
        .CheckBoxRunFlats.Value = False
        ' ' falty section ' '
        
        
        ' jak wypelnic arkusz puses
        ' ============================
        
        .OptionButtonPUSMGO.Value = False
        .OptionButtonPUSMIXED.Value = False
        .OptionButtonPUSWIZARD.Value = False
        
        ' ============================

        '' cbal section ''
        ' ====================================
        .OptionButtonCbalFromMGO.Value = False
        .OptionButtonCbalFromWGEN.Value = False
        .OptionButtonCbalFromWizard.Value = False
        
        ' ====================================
        
        ' recv section
        ' ====================================
        .CheckBoxRECV.Value = False
        ' ====================================
        
        
        
        
        ' coverage and coord lilst
        ' ====================================
        .CheckBoxRunCov.Value = False
        .CheckBoxCoordList.Value = False
        ' ====================================
        
        ' cov and coord list run type
        ' ====================================
        .OptionButtonPUSesFromMGO.Value = False
        .OptionButtonPUSesFromWiz.Value = False
        
        
        ' ====================================
        
        
        ' ====================================
        .connectWithExternalIch ich
        .Show
        
    End With
    
    
    Set ich = Nothing
    
    
    MsgBox "ready!"
End Sub


