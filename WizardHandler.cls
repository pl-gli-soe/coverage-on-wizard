VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WizardHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IPUSTaker

Private w As Workbook
Private psh As Worksheet

Private r As Range

Private ph As PusHandler


Private Sub m_pobierz_pusy(s As String)

    Set w = Application.Workbooks(CStr(s))
    ' to jest href do zew polaczenia
    


    If sprawdz_czy_wybrany_sheet_to_sheet_pusow_w_wizardzie() Then
    
        ' href do maina - super side effect ale coz
        ' ww jest linkiem przez WizardHandler,
        ' potem przez formularz wybierania pliku
        ' i na koniec w glownej klasie CowH
        ' gdzie tmp w foo take_data_from_wizard
        ' jest zwracany do pola source_workbook
        ' troche masakra - nawet nie wiem czy bedzie to dzialac.
        
        ' problem rozwiazalem zmienna globalna
        ' source_wizard - nie jest to najpiekniejsze rozwiazanie
        ' ale taki szeroki cross byl dosyc upierdliwy
        ' ====================================================================
        ' ====================================================================
        ' ====================================================================
        Set G_SOURCE_WIZARD = w
        ' ====================================================================
        ' ====================================================================
        ' ====================================================================
    
        Set psh = w.Sheets(Cow.PICKUPS_SH_NM)
        
        
        Set ph = New PusHandler
        ph.take_puses_from_wizard psh
        ph.refresh_puses_sheet
        ph.put_puses_into_this_puses_sh
        Set ph = Nothing
    Else
        MsgBox "to nie jest plik typu Wizard!"
        End
    End If

        
End Sub

Private Function sprawdz_czy_wybrany_sheet_to_sheet_pusow_w_wizardzie() As Boolean
    sprawdz_czy_wybrany_sheet_to_sheet_pusow_w_wizardzie = False
    
    
    
    Set psh = Nothing
    On Error Resume Next
    Set psh = w.Sheets(Cow.PICKUPS_SH_NM)
    
    
    If psh Is Nothing Then
        sprawdz_czy_wybrany_sheet_to_sheet_pusow_w_wizardzie = False
    Else
        sprawdz_czy_wybrany_sheet_to_sheet_pusow_w_wizardzie = True
    End If
    
    
End Function


Private Sub IPUSTaker_pobierzPusy(mich As InitConfigHandler, Optional s As String)
    If mich Is Nothing Then
        m_pobierz_pusy CStr(s)
    Else
        MsgBox "zla konfiguracja suba pobierz pusy"
    End If
End Sub
