VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CfgUserForm 
   Caption         =   "Init Config Form"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7785
   OleObjectBlob   =   "CfgUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CfgUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2015 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Private ich As InitConfigHandler


Public Sub connectWithExternalIch(mich As InitConfigHandler)
    Set ich = mich
End Sub

Private Sub BtnSubmit_Click()
    hide
    
    ' tutaj pozniej, jak dopisze implementacje
    ' nalezy wsadzic faktycznie istniejace warunki
    ' ich.cbalFromWhere = E_CBAL_FROM_MGO
    
    
    If Me.OptionButtonBalanceFromCBAL.Value = True Then
        ich.coverageStockBasedOnQuestion = E_STOCK_FROM_CBAL
    ElseIf Me.OptionButtonBalanceFromTotalMRDQty.Value = True Then
        ich.coverageStockBasedOnQuestion = E_STOCK_FROM_TOTAL_MRD_QTY
    End If
    
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
    
    
    If Len(Me.TextBoxFUPCODEFilter.Value) <> 2 Then
        Me.TextBoxFUPCODEFilter.Value = ""
    End If
    
    If Trim(Me.TextBoxFUPCODEFilter.Value) = "" Then
        Me.TextBoxFUPCODEFilter.Value = ""
    End If
    
    ich.fup_code = CStr(Me.TextBoxFUPCODEFilter.Value)
    G_FUP_CODE = ich.fup_code
    
    
    main ich
End Sub

Private Sub CheckBoxRunCov_Change()


    'With Me
    '    If .CheckBoxRunCov.Value = True Then
    '        .CheckBoxRunFlats.Value = True
    '
    '        zmien_flaty True
    '    Else
    '        .CheckBoxRunFlats.Value = False
    '
    '        zmien_flaty False
    '    End If
    'End With

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

