VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CfgUserForm 
   Caption         =   "Init Config Form"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7770
   OleObjectBlob   =   "CfgUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CfgUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ich As InitConfigHandler


Public Sub connectWithExternalIch(mich As InitConfigHandler)
    Set ich = mich
End Sub

Private Sub BtnSubmit_Click()
    hide
    
    ' tutaj pozniej, jak dopisze implementacje
    ' nalezy wsadzic faktycznie istniejace warunki
    ' ich.cbalFromWhere = E_CBAL_FROM_MGO
    
    If Me.OptionButtonPUSMIXED.Value = True Then
        ich.pusFromWhere = E_PUS_MIX
    ElseIf Me.OptionButtonPUSMGO.Value = True Then
        ich.pusFromWhere = E_PUS_MGO
    ElseIf Me.OptionButtonPUSWIZARD.Value = True Then
        ich.pusFromWhere = E_PUS_WIZARD
    End If
    
    
    ' teraz jakie
    ich.pusFlatTable = Me.CheckBoxPUSes.Value
    ich.rqmFlatTable = Me.CheckBoxRQMs.Value
    
    ich.addRecv = Me.CheckBoxRECV.Value
    
    
    If Me.OptionButtonCbalFromWizard.Value = True Then
        ich.cbalFromWhere = E_CBAL_FROM_WIZARD
    ElseIf Me.OptionButtonCbalFromMGO.Value = True Then
        ich.cbalFromWhere = E_CBAL_FROM_MGO
    End If
    
    ich.cbalFlatTable = Me.CheckBoxCBALs.Value
    
    ' esy
    ich.flats = Me.CheckBoxRunFlats.Value
    ich.do_we_want_to_run_coverage = Me.CheckBoxRunCov.Value
    ich.do_we_want_to_run_coord_list = Me.CheckBoxCoordList.Value
    
    If Me.OptionButtonPUSesFromMGO.Value = True Then
        ich.pusesForCoverage = E_TYPE_PUS_MGO
    ElseIf Me.OptionButtonPUSesFromMGO.Value = True Then
        ich.pusesForCoverage = E_TYPE_PUS_WIZARD
    End If
    
    
    main ich
End Sub

Private Sub CheckBoxRunCov_Change()


    With Me
        If .CheckBoxRunCov.Value = True Then
            .CheckBoxRunFlats.Value = True
            
            zmien_flaty True
        Else
            .CheckBoxRunFlats.Value = False
            
            zmien_flaty False
        End If
    End With

End Sub




Private Sub zmien_flaty(arg As Boolean)
    
    With Me
        .CheckBoxPUSes.Value = arg
        .CheckBoxRQMs.Value = arg
        .CheckBoxCBALs.Value = arg
    End With
End Sub

Private Sub CheckBoxRunFlats_Change()

    If Me.CheckBoxRunFlats.Value = True Then
        zmien_flaty True
    Else
        zmien_flaty False
    End If

End Sub

