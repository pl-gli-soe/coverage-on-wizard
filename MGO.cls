VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
Private Sessions As Object
Private System As Object
Private sess0 As Object


Private scr_ms9pop00 As MS9POP00
Private scr_ms9po400 As MS9PO400
Private scr_zk7pwrqm As ZK7PWRQM
Private scr_zk7pdrqm As ZK7PDRQM
Private scr_zk7ppus0 As ZK7PPUS0
Private scr_ms9ph100 As MS9PH100
Private scr_ms7p5200 As MS7P5200
Private scr_zk7ppum0 As ZK7PPUM0
Private scr_zk7pprul As ZK7PPRUL
Private scr_zk7pprhr As ZK7PPRHR
Private scr_zk7ppud0 As ZK7PPUD0
Private scr_ms3p9800 As MS3P9800



Private Sub Class_Initialize()
    Set System = CreateObject("EXTRA.System")   ' Gets the system object
    Set Sessions = System.Sessions
    Set sess0 = System.activeSession
    
    Set scr_ms9pop00 = New MS9POP00
    scr_ms9pop00.sesja = sess0
    waitForMgo
    
    Set scr_ms9po400 = New MS9PO400
    scr_ms9po400.sesja = sess0
    waitForMgo
    
    Set scr_zk7pwrqm = New ZK7PWRQM
    scr_zk7pwrqm.sesja = sess0
    waitForMgo
    
    Set scr_zk7pdrqm = New ZK7PDRQM
    scr_zk7pdrqm.sesja = sess0
    waitForMgo
    
    Set scr_zk7ppus0 = New ZK7PPUS0
    scr_zk7ppus0.sesja = sess0
    waitForMgo
    
    Set scr_ms9ph100 = New MS9PH100
    scr_ms9ph100.sesja = sess0
    waitForMgo
    
    Set scr_ms7p5200 = New MS7P5200
    scr_ms7p5200.sesja = sess0
    waitForMgo
    
    Set scr_zk7ppum0 = New ZK7PPUM0
    scr_zk7ppum0.sesja = sess0
    waitForMgo
    
    Set scr_zk7pprul = New ZK7PPRUL
    scr_zk7pprul.sesja = sess0
    waitForMgo
    
    Set scr_zk7pprhr = New ZK7PPRHR
    scr_zk7pprhr.sesja = sess0
    waitForMgo
    
    Set scr_zk7ppud0 = New ZK7PPUD0
    scr_zk7ppud0.sesja = sess0
    waitForMgo
    
    Set scr_ms3p9800 = New MS3P9800
    scr_ms3p9800.sesja = sess0
    waitForMgo
    
End Sub

Private Sub Class_Terminate()
    Set System = Nothing
    Set Sessions = Nothing
    Set sess0 = Nothing
End Sub

Public Property Get actualScreen() As String
    actualScreen = sess0.screen.getString(1, 2, 8)
End Property

Public Property Get activeSession() As Object
    Set activeSession = sess0
End Property


Public Function convertToDateFromMS9PO400Date(ms9po400_date As String) As Date


    convertToDateFromMS9PO400Date = Date

    ' format daty dla scr MS9PO400 -    ' 27MR12
    ' -----------------------------------------------------------YEAR----------------------------------------MONTH-------------------------------------------------------YEAR-------------------------------
    convertToDateFromMS9PO400Date = CDate(CStr(2000 + Int(Right(ms9po400_date, 2))) & "-" & CStr(FromMGOMonth(Right(Left(ms9po400_date, 4), 2))) & "-" & CStr(Int(Left(ms9po400_date, 2))))
End Function

Public Function convertToDateFromMS9POP00Date(ms9po400_date As String) As Date


    convertToDateFromMS9POP00Date = Date

    ' format daty dla scr MS9POP00 -    ' 27MR12
    ' -----------------------------------------------------------YEAR----------------------------------------MONTH-------------------------------------------------------YEAR-------------------------------
    convertToDateFromMS9POP00Date = CDate(CStr(2000 + Int(Right(ms9po400_date, 2))) & "-" & CStr(FromMGOMonth(Right(Left(ms9po400_date, 5), 2))) & "-" & CStr(Int(Left(ms9po400_date, 2))))
End Function


Public Function convertToDateFromMS9PH100ShortDate(ms9ph100_shortdate As String) As Date
    Dim tmp_date As Date
    tmp_date = CDate(Format(Now, "yyyy-mm-dd")) - 100

    If tmp_date <= CDate(CStr(Format(Now, "yyyy")) & "-" & CStr(FromMGOMonth(Right(ms9ph100_shortdate, 2))) & "-" & CStr(Left(ms9ph100_shortdate, 2))) Then
        
        '--------------------------------------------------YEAR-------------------------------------MONTH-------------------------------------DAY--------------
        convertToDateFromMS9PH100ShortDate = CDate(CStr(Format(Now, "yyyy")) & "-" & CStr(FromMGOMonth(Right(ms9ph100_shortdate, 2))) & "-" & CStr(Left(ms9ph100_shortdate, 2)))
    Else
        '--------------------------------------------------YEAR-------------------------------------MONTH---------------------------------------------DAY--------------
        convertToDateFromMS9PH100ShortDate = CDate(CStr(Int(Format(Now, "yyyy") + 1)) & "-" & CStr(FromMGOMonth(Right(ms9ph100_shortdate, 2))) & "-" & CStr(Left(ms9ph100_shortdate, 2)))
    End If
End Function

Public Function convertToDateFromMS9POP00ShortDate(ms9pop00_shortdate As String) As Date


    ' instrukcja warunkowa wynika z powodu braku roku na ekranie zk7pdrqm i przez to program musi sam sie domyslec
    ' ze faktycznie nastapila zmiana roku- sam rozwiazalem to tak ze zazwyczaj dzieje sie tak nie dobrze, gdy chcemy
    ' stworzyc teardown nachodzi to na nowy rok, ale my jako ludzie zobaczymy tylko i wylacznie ze zaczynamy z powrotem
    ' od stycznia... postanowi�em zatem zostawic na stale ten sam rok i w przyapdku gdy przejdziemy do stycznia
    ' i sprawdzimy warunkiem to tak jakbysmy sie cofneli w czasie- cofniecie w czasie algorytm potraktuje jako co� osobliwego
    ' zgodnie z trywialnym zalozeniem doda plus jeden do aktualnego roku
    '
    '
    '
    ' na dzien 4 maja 2012
    
    
    ' JN03  - example


    '' Debug.Print CStr(Format(Now, "yyyy"))
    '' Debug.Print CStr(FromMGOMonth(Right(zk7pdrqm_date, 2)))
    '' Debug.Print CStr(Left(zk7pdrqm_date, 2))
    '
    '
    '' Debug.Print CStr(CDate(CStr(Format(Now, "yyyy")) & "-" & CStr(FromMGOMonth(Right(zk7pdrqm_date, 2))) & "-" & CStr(Left(zk7pdrqm_date, 2))))
 
    ' tutaj mamy troche malo pieknie gdzie zasieg ff jest na tyle maly ze 100 dni jest out of scope dzieki czemu wie ze przeskakujemy na nastepny dzien
    ' jesli miedzy poszczegolnymi dniami iteracji jest az taka rozbieznosc moze sie zdarzyc ze bedzie 200 down days to wtedy fireflake
    ' sie wysypie faktycznie :P
    Dim tmp_date As Date
    tmp_date = CDate(Format(Now, "yyyy-mm-dd")) - 100

    If tmp_date <= CDate(CStr(Format(Now, "yyyy")) & "-" & CStr(FromMGOMonth(Left(ms9pop00_shortdate, 2))) & "-" & CStr(Right(ms9pop00_shortdate, 2))) Then
        
        '--------------------------------------------------YEAR-------------------------------------MONTH-------------------------------------DAY--------------
        convertToDateFromMS9POP00ShortDate = CDate(CStr(Format(Now, "yyyy")) & "-" & CStr(FromMGOMonth(Left(ms9pop00_shortdate, 2))) & "-" & CStr(Right(ms9pop00_shortdate, 2)))
    Else
        '--------------------------------------------------YEAR-------------------------------------MONTH---------------------------------------------DAY--------------
        convertToDateFromMS9POP00ShortDate = CDate(CStr(Int(Format(Now, "yyyy") + 1)) & "-" & CStr(FromMGOMonth(Left(ms9pop00_shortdate, 2))) & "-" & CStr(Right(ms9pop00_shortdate, 2)))
    End If
End Function



Public Function convertToDateFromZK7PDRQMDate(zk7pdrqm_date As String) As Date
    
    ' instrukcja warunkowa wynika z powodu braku roku na ekranie zk7pdrqm i przez to program musi sam sie domyslec
    ' ze faktycznie nastapila zmiana roku- sam rozwiazalem to tak ze zazwyczaj dzieje sie tak nie dobrze, gdy chcemy
    ' stworzyc teardown nachodzi to na nowy rok, ale my jako ludzie zobaczymy tylko i wylacznie ze zaczynamy z powrotem
    ' od stycznia... postanowi�em zatem zostawic na stale ten sam rok i w przyapdku gdy przejdziemy do stycznia
    ' i sprawdzimy warunkiem to tak jakbysmy sie cofneli w czasie- cofniecie w czasie algorytm potraktuje jako co� osobliwego
    ' zgodnie z trywialnym zalozeniem doda plus jeden do aktualnego roku
    '
    '
    '
    ' na dzien 4 maja 2012
    
    
    ' 03.JN - example


    '' Debug.Print CStr(Format(Now, "yyyy"))
    '' Debug.Print CStr(FromMGOMonth(Right(zk7pdrqm_date, 2)))
    '' Debug.Print CStr(Left(zk7pdrqm_date, 2))
    '
    '
    '' Debug.Print CStr(CDate(CStr(Format(Now, "yyyy")) & "-" & CStr(FromMGOMonth(Right(zk7pdrqm_date, 2))) & "-" & CStr(Left(zk7pdrqm_date, 2))))
 
    ' tutaj mamy troche malo pieknie gdzie zasieg ff jest na tyle maly ze 100 dni jest out of scope dzieki czemu wie ze przeskakujemy na nastepny dzien
    ' jesli miedzy poszczegolnymi dniami iteracji jest az taka rozbieznosc moze sie zdarzyc ze bedzie 200 down days to wtedy fireflake
    ' sie wysypie faktycznie :P
    Dim tmp_date As Date
    tmp_date = CDate(Format(Now, "yyyy-mm-dd")) - 100

    If tmp_date <= CDate(CStr(Format(Now, "yyyy")) & "-" & CStr(FromMGOMonth(Right(zk7pdrqm_date, 2))) & "-" & CStr(Left(zk7pdrqm_date, 2))) Then
        '' Debug.Print CStr(Format(Now, "yyyy"))
        '' Debug.Print CStr(FromMGOMonth(Right(zk7pdrqm_date, 2)))
        '' Debug.Print CStr(Left(zk7pdrqm_date, 2))
        
        '--------------------------------------------------YEAR-------------------------------------MONTH-------------------------------------DAY--------------
        convertToDateFromZK7PDRQMDate = CDate(CStr(Format(Now, "yyyy")) & "-" & CStr(FromMGOMonth(Right(zk7pdrqm_date, 2))) & "-" & CStr(Left(zk7pdrqm_date, 2)))
    Else
        '--------------------------------------------------YEAR-------------------------------------MONTH---------------------------------------------DAY--------------
        convertToDateFromZK7PDRQMDate = CDate(CStr(Int(Format(Now, "yyyy") + 1)) & "-" & CStr(FromMGOMonth(Right(zk7pdrqm_date, 2))) & "-" & CStr(Left(zk7pdrqm_date, 2)))
    End If
End Function

Public Function convertToDateFromZK7PPUS0Date(zk7ppus0_date As String) As Date
    Dim d As String
    d = Right(CStr(zk7ppus0_date), 4) & "-" & CStr(FromMGOMonth(Right(Left(CStr(zk7ppus0_date), 5), 2))) & "-" & Left(CStr(zk7ppus0_date), 2)
    convertToDateFromZK7PPUS0Date = CDate(Format(d, "yyyy-mm-dd"))
End Function

Public Function convertToDateFromZK7PPUM0Date(zk7ppum0_date As String) As Date
    Dim d As String
    d = CStr(Left(CStr(Format(Now, "yyyy")), 2)) & Right(CStr(zk7ppum0_date), 2) & "-" & CStr(FromMGOMonth(Right(Left(CStr(zk7ppum0_date), 5), 2))) & "-" & Left(CStr(zk7ppum0_date), 2)
    convertToDateFromZK7PPUM0Date = CDate(Format(d, "yyyy-mm-dd"))
End Function

Public Function convertToMgoLongDate(n_date As Date) As String
    convertToMgoLongDate = Format(n_date, "dd") + "." + MgoMonth(n_date) + "." + Format(n_date, "yyyy")
End Function

Public Function convertToMgoShortDate(n_date As Date) As String
    convertToMgoShortDate = Format(n_date, "dd") + "." + MgoMonth(n_date) + "." + Format(n_date, "yy")
End Function

Public Function FromMGOMonth(mgo_month As String)

    Select Case mgo_month
        Case "JA"
            FromMGOMonth = 1
        Case "FE"
            FromMGOMonth = 2
        Case "MR"
            FromMGOMonth = 3
        Case "AP"
            FromMGOMonth = 4
        Case "MY"
            FromMGOMonth = 5
        Case "JN"
            FromMGOMonth = 6
        Case "JL"
            FromMGOMonth = 7
        Case "AU"
            FromMGOMonth = 8
        Case "SE"
            FromMGOMonth = 9
        Case "OC"
            FromMGOMonth = 10
        Case "NO"
            FromMGOMonth = 11
        Case "DE"
            FromMGOMonth = 12
        Case Else
            FromMGOMonth = -1
    End Select
End Function

Public Function getString(x As Integer, y As Integer, size As Integer) As String
    getString = sess0.screen.getString(x, y, size)
    waitForMgo
End Function

Public Sub mgoInit()
    Set System = CreateObject("EXTRA.System")   ' Gets the system object
    Set Sessions = System.Sessions
    Set sess0 = System.activeSession
    
    waitForMgo
End Sub

Public Function MgoMonth(n_date As Date) As String
    Select Case Format(n_date, "m") - 1
        Case 0
            MgoMonth = "JA"
        Case 1
            MgoMonth = "FE"
        Case 2
            MgoMonth = "MR"
        Case 3
            MgoMonth = "AP"
        Case 4
            MgoMonth = "MY"
        Case 5
            MgoMonth = "JN"
        Case 6
            MgoMonth = "JL"
        Case 7
            MgoMonth = "AU"
        Case 8
            MgoMonth = "SE"
        Case 9
            MgoMonth = "OC"
        Case 10
            MgoMonth = "NO"
        Case 11
            MgoMonth = "DE"
    End Select
End Function


Public Property Let pMS3P9800(RHS As MS3P9800)
    Set scr_ms3p9800 = RHS
End Property

Public Property Get pMS3P9800() As MS3P9800
    Set pMS3P9800 = scr_ms3p9800
End Property



Public Property Let pMS7P5200(RHS As MS7P5200)
    Set scr_ms7p5200 = RHS
End Property

Public Property Get pMS7P5200() As MS7P5200
    Set pMS7P5200 = scr_ms7p5200
End Property

Public Property Let pMS9PH100(RHS As MS9PH100)
    Set scr_ms9ph100 = RHS
End Property

Public Property Get pMS9PH100() As MS9PH100
    Set pMS9PH100 = scr_ms9ph100
End Property

Public Property Let pMS9PO400(RHS As MS9PO400)
    Set scr_ms9po400 = RHS
End Property

Public Property Get pMS9PO400() As MS9PO400
    Set pMS9PO400 = scr_ms9po400
End Property

Public Property Let pMS9POP00(RHS As MS9POP00)
    Set scr_ms9pop00 = RHS
End Property

Public Property Get pMS9POP00() As MS9POP00
    Set pMS9POP00 = scr_ms9pop00
End Property

Public Sub putString(str As String, x As Integer, y As Integer)
    sess0.screen.putString str, x, y
    waitForMgo
End Sub

Public Property Let pZK7PDRQM(RHS As ZK7PDRQM)
    Set scr_zk7pdrqm = RHS
End Property

Public Property Get pZK7PDRQM() As ZK7PDRQM
    Set pZK7PDRQM = scr_zk7pdrqm
End Property

Public Property Let pZK7PPUS0(RHS As ZK7PPUS0)
    Set scr_zk7ppus0 = RHS
End Property

Public Property Get pZK7PPUS0() As ZK7PPUS0
    Set pZK7PPUS0 = scr_zk7ppus0
End Property

Public Property Let pZK7PWRQM(RHS As ZK7PWRQM)
    Set scr_zk7pwrqm = RHS
End Property

Public Property Get pZK7PWRQM() As ZK7PWRQM
    Set pZK7PWRQM = scr_zk7pwrqm
End Property

Public Property Let pZK7PPUM0(RHS As ZK7PPUM0)
    Set scr_zk7ppum0 = RHS
End Property

Public Property Get pZK7PPUM0() As ZK7PPUM0
    Set pZK7PPUM0 = scr_zk7ppum0
End Property

Public Property Let pZK7PPRUL(RHS As ZK7PPRUL)
    Set scr_zk7pprul = RHS
End Property

Public Property Get pZK7PPRUL() As ZK7PPRUL
    Set pZK7PPRUL = scr_zk7pprul
End Property

Public Property Let pZK7PPRHR(RHS As ZK7PPRHR)
    Set scr_zk7pprhr = RHS
End Property

Public Property Get pZK7PPRHR() As ZK7PPRHR
    Set pZK7PPRHR = scr_zk7pprhr
End Property


Public Property Let pZK7PPUD0(RHS As ZK7PPUD0)
    Set scr_zk7ppud0 = RHS
End Property

Public Property Get pZK7PPUD0() As ZK7PPUD0
    Set pZK7PPUD0 = scr_zk7ppud0
End Property


Public Sub submit()
    sess0.screen.sendKeys ("<Enter>")
    waitForMgo
End Sub

Public Sub sendKeys(str As String)
    sess0.screen.sendKeys (CStr(str))
    waitForMgo
End Sub

Public Function getCode(mgo_screen) As String

    If mgo_screen = "zk7pdrqm" Or mgo_screen = "ms9po400" Or mgo_screen = "zk7ppum0" Or mgo_screen = "ms9ph100" Then
        getCode = CStr(getString(22, 2, 5))
    ElseIf mgo_screen = "ms9pop00" Then
        getCode = CStr(getString(23, 2, 5))
    End If
    waitForMgo
End Function

Public Sub waitForMgo()
    Do
        DoEvents
    Loop Until sess0.screen.OIA.XStatus = 0
End Sub
