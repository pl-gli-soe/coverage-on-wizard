VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PN_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' pn jako key
Public pn_dic As Dictionary
Public plt As String

Private psh As Worksheet
Private ish As Worksheet




Public Sub skompletuj_pny_z_arkusza_puses()


    

    Dim rr As Range
    Set rr = psh.Cells(2, Cow.PUSES_PN)
    
    Do
        If Trim(rr) <> "" Then
            If pn_dic.Exists(Trim(rr)) Then
            Else
                pn_dic.Add Trim(rr), 1
            End If
        End If
        Set rr = rr.Offset(1, 0)
    Loop Until Trim(rr) = ""

End Sub

Public Sub dodaj_pny_z_arkusza_input()
    
    Dim rr As Range
    Set rr = ish.Cells(2, 1)
    
    
    Do
        If Trim(rr) <> "" Then
            If pn_dic.Exists(Trim(rr)) Then
            Else
                If Trim(rr) <> "" Then
                    pn_dic.Add Trim(rr), 1
                End If
            End If
        End If
        Set rr = rr.Offset(1, 0)
    Loop Until Trim(rr) = ""

End Sub

Public Sub pobierz_plt_raz()

    ' dane powinny byc zawarte w tym arkuszu po zebraniu danych z wizarda
    plt = ThisWorkbook.Sheets(Cow.PUSES_SH_NM).Range("A2")
    
    If Trim(plt) = "" Then
        ' nie bylo danych z wizarda zatem trzeba sciagnac plt z arkusza input
        If Trim(ThisWorkbook.Sheets(Cow.INPUT_SH_NM).Range("plt")) <> "" Then
            plt = ThisWorkbook.Sheets(Cow.INPUT_SH_NM).Range("plt")
        End If
    End If
    
    If Trim(plt) = "" Then
        MsgBox "cos poszlo nie tak podczas proby zaciagniecia danych z mgo"
        End
    End If

End Sub

Private Sub Class_Initialize()
    Set pn_dic = New Dictionary
    Set psh = ThisWorkbook.Sheets(Cow.PUSES_SH_NM)
    Set ish = ThisWorkbook.Sheets(Cow.INPUT_SH_NM)
End Sub

Private Sub Class_Terminate()
    Set pn_dic = Nothing
End Sub
