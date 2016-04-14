VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CowHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mich As InitConfigHandler
Private thissh As Worksheet


Private master_sheet_from_source As Worksheet
Private details_sheet_from_source As Worksheet


Public Sub init(ich As InitConfigHandler)
    Set mich = ich
    run
End Sub


Private Sub Class_Initialize()
    Set mich = Nothing
    Set G_SOURCE_WIZARD = Nothing
End Sub

Private Sub Class_Terminate()
    Set mich = Nothing
    Set G_SOURCE_WIZARD = Nothing
End Sub


Private Sub run()
    
    If mich.pusFlatTable Then
        create_pus_flat_table
    End If
    
    If mich.rqmFlatTable Then
        create_rqm_flat_table
    End If
    
    If mich.cbalFlatTable Then
        create_cbal_flat_table
    End If
    
    If mich.do_we_want_to_run_coverage Then
        create_coverage
    End If
    
    If mich.do_we_want_to_run_coord_list Then
        create_coord_list
    End If
End Sub


Private Sub create_cbal_flat_table()
    refresh_sheet CStr(Cow.CBALS_SH_NM)
    
    If mich.cbalFromWhere = E_CBAL_FROM_MGO Then
    
        take_cbal_from_mgo
    ElseIf mich.cbalFromWhere = E_CBAL_FROM_WIZARD Then
        take_cbal_from_wizard
    Else
        MsgBox "wciaz brak odpowiedniej implementacji"
    End If
End Sub

Private Sub create_pus_flat_table()

    refresh_sheet CStr(Cow.PUSES_SH_NM)

    If mich.pusFromWhere = E_PUS_MIX Then
    
        take_data_from_wizard
        ' musi bazowac na tym co pokazal Wizard (taka opcja: MIX)
        ' MsgBox source_workbook.Name
        take_data_from_mgo
        
        
        
        
    Else
        If mich.pusFromWhere = E_PUS_MGO Then
            take_data_from_mgo
        ElseIf mich.pusFromWhere = E_PUS_WIZARD Then
            
            take_data_from_wizard
        End If
    End If
    
    autofit_sheet
    sort_puses_sheet
    wrap_text_on_puses
    colour_same_pns
    

End Sub

Private Sub autofit_sheet()

    With thissh
        
        .Columns("A:ZZ").AutoFit
    End With
End Sub

Private Sub sort_puses_sheet()
    ThisWorkbook.Sheets("PUSes").Sort.SortFields.Clear
    ThisWorkbook.Sheets("PUSes").Sort.SortFields.Add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ThisWorkbook.Sheets("PUSes").Sort
        .SetRange Range("A2:O100000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Sub wrap_text_on_puses()
    With ThisWorkbook.Sheets("PUSes").Range("A2:O100000")
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Private Sub colour_same_pns()
    
    ' tbd
    ' ---------------------------------
    
    
    ' ---------------------------------
End Sub

Private Sub refresh_sheet(shnm As String)

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(CStr(shnm)).Delete
    Application.DisplayAlerts = True
    
    Set thissh = ThisWorkbook.Sheets.Add
    On Error Resume Next
    thissh.Name = shnm

    If shnm = Cow.PUSES_SH_NM Then
    
        ' po refreshu jest juz zlapany
        'this_psh.Cells(x, 1) = p.plt
        'this_psh.Cells(x, 2) = item.pn
        'this_psh.Cells(x, 3) = p.fup_code
        'this_psh.Cells(x, 4) = p.sdate
        'this_psh.Cells(x, 5) = p.eda
        'this_psh.Cells(x, 6) = item.qty
        'this_psh.Cells(x, 7) = item.del_qty
        'this_psh.Cells(x, 8) = item.recv
        'this_psh.Cells(x, 9) = p.recv
        'this_psh.Cells(x, 10) = p.pusname
        'this_psh.Cells(x, 11) = p.duns
        'this_psh.Cells(x, 12) = p.supp_nm
        'this_psh.Cells(x, 13) = p.on_mgo
        'this_psh.Cells(x, 14) = p.on_wizard
        thissh.Cells(1, 1) = "PLT"
        thissh.Cells(1, 2) = "PN"
        thissh.Cells(1, 3) = "FUP_CODE"
        thissh.Cells(1, 4) = "PUS_DATE"
        thissh.Cells(1, 5) = "DEL_DATE"
        thissh.Cells(1, 6) = "WIZ QTY"
        thissh.Cells(1, 7) = "MGO QTY"
        thissh.Cells(1, 8) = "RECV"
        thissh.Cells(1, 9) = "RECV STATUS"
        thissh.Cells(1, 10) = "PUS_NAME"
        thissh.Cells(1, 11) = "DUNS"
        thissh.Cells(1, 12) = "SUPPLIER NAME"
        thissh.Cells(1, 13) = "ON MGO"
        thissh.Cells(1, 14) = "ON WIZARD"
        thissh.Cells(1, 15) = "LOG"
        
    ElseIf shnm = Cow.RQMS_SH_NM Then
        
        thissh.Cells(1, 1) = "PLT"
        thissh.Cells(1, 2) = "PN"
        thissh.Cells(1, 3) = "FUP_CODE"
        thissh.Cells(1, 4) = "CW"
        thissh.Cells(1, 5) = "QTY"
        
    ElseIf shnm = Cow.CBALS_SH_NM Then
        
        thissh.Cells(1, 1) = "PLT"
        thissh.Cells(1, 2) = "PN"
        thissh.Cells(1, 3) = "CBAL"
    End If
End Sub

Private Sub take_data_from_wizard()
    With WybierzPlikForm
        .init FOMULARZ_WYBORU_PLIKU_DLA_PUS
        .Show
    End With
    
End Sub

Private Sub take_cbal_from_mgo()
    
    Dim rqmsy_z_mgo As ICBalFromHandler
    Set rqmsy_z_mgo = New CBalFromMGOHandler
    
    rqmsy_z_mgo.pobierzCbale mich
End Sub

Private Sub take_data_from_mgo()


    If mich.pusFromWhere = E_PUS_MIX Or mich.pusFromWhere = E_PUS_MGO Then
    
    
        ' to jest arkusz w ktorym juz znajduja sie dane z wizarda
        ' thissh
        ' ========================================================================
        Dim pusy_z_mgo As IPUSTaker
        Set pusy_z_mgo = New PusFromMGOHandler
        
        pusy_z_mgo.pobierzPusy mich
        ' ========================================================================
    Else
        MsgBox "brak obslugi dla tej implementacji"
    End If
    
    
    
End Sub


Private Sub take_cbal_from_wizard()
    
    refresh_sheet CStr(Cow.CBALS_SH_NM)
    
    If G_SOURCE_WIZARD Is Nothing Then
        
        With WybierzPlikForm
        .init FOMULARZ_WYBORU_PLIKU_DLA_RQM
        .Show
    End With
    Else
        ' prostszy scenario poniewaz juz mamy polaczenie miedzy plikami
        ' -------------------------------------------------------------
        
        'Private master_sheet_from_source As Worksheet
        'Private details_sheet_from_source As Worksheet
        
        
        Dim c As ICBalFromHandler
        Set c = New CBalFromWizardHandler
        c.pobierzCbale mich
        
        
        
        ' -------------------------------------------------------------
    End If
End Sub


Private Sub create_rqm_flat_table()


    refresh_sheet CStr(Cow.RQMS_SH_NM)

    If G_SOURCE_WIZARD Is Nothing Then
        
        With WybierzPlikForm
        .init FOMULARZ_WYBORU_PLIKU_DLA_RQM
        .Show
    End With
    Else
        ' prostszy scenario poniewaz juz mamy polaczenie miedzy plikami
        ' -------------------------------------------------------------
        
        'Private master_sheet_from_source As Worksheet
        'Private details_sheet_from_source As Worksheet
        
        
        Dim r As IRqmTaker
        Set r = New RqmsFromWizard
        r.pobierzRqmsy mich, G_SOURCE_WIZARD
        
        
        
        ' -------------------------------------------------------------
    End If
End Sub

Private Sub create_coverage()

    If check_flats() Then
        
        ' jest OK
        ' ----------------------------------------
        
        Dim c As ICoverage
        Set c = New Coverage
        
        c.makeCoverage mich
        
        Set c = Nothing
        ' ----------------------------------------
    Else
        ' cos nie tak z flatami
        ' ========================================
        MsgBox "cos nie tak z flatami!"
        End
        ' ========================================
    End If
End Sub

Private Sub create_coord_list()
    
    If check_(Cow.PUSES_SH_NM) And check_(Cow.CBALS_SH_NM) Then
        Dim c As ICoverage
        
        
        Set c = New CoordList
        c.makeCoverage mich
    Else
        ' cos nie tak z flatami
        ' ========================================
        MsgBox "cos nie tak z flatami!"
        End
        ' ========================================
    End If
End Sub


Private Function check_flats() As Boolean
    check_flats = False
    
    
    ' ale jest to racze prymitywny check tylko sprawdzajacy czy w ogole jest jaki kolwiek content
    
    If check_(Cow.PUSES_SH_NM) And check_(Cow.RQMS_SH_NM) And check_(Cow.CBALS_SH_NM) Then
        check_flats = True
    End If
End Function

Private Function check_(s)

    check_ = False

    Dim sh As Worksheet, r As Range
    Set sh = ThisWorkbook.Sheets(CStr(s))
    Set r = sh.Range(sh.Cells(2, 1), sh.Cells(1000, 100))
    
    If r.COUNT > 1 Then
        check_ = True
    End If
    
    
End Function

