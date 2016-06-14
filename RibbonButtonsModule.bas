Attribute VB_Name = "RibbonButtonsModule"
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
        
        
        ' sekcja stocku
        ' ====================================
        '.OptionButtonBalanceFromCBAL.Value = False
        '.OptionButtonBalanceFromTotalMRDQty.Value = False
        .OptionButtonBalanceOnZero.Value = True
        
        ' ====================================
        
        ' fup code section
        ' ====================================
        .TextBoxFUPCODEFilter.Value = ""
        
        
        ' ====================================
        .connectWithExternalIch ich
        .show
        
    End With
    
    
    Set ich = Nothing
    
    
    MsgBox "ready!"
End Sub


