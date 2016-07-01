Attribute VB_Name = "VersionModule"
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


' ver 0.x
' -----------------------------------------------------------------------------------------------------------------

' ver 0.1
' ==================================================================================================
'
' init on this project
' for now only version module and export module
' later on will provide MGO handler from fire flake light
'
' ==================================================================================================


' ver 0.2
' ==================================================================================================
'
' copy fetching process from wizard macro
'
' READY:
' flat PUS from Wizard (pierwsze kolumny)
' TO DO:
' duns plt supp nm from wizard jeszcze
'
' ==================================================================================================


' ver 0.3
' ==================================================================================================
'
' copy fetching process from wizard macro
'
' READY:
' flat PUS from Wizard (pierwsze kolumny)
' TO DO:
' duns plt supp nm from wizard jeszcze

' wstepne polaczenie pusow z wizarda z pusami z mgo
' pojawil sie problem z numerami pta ktore nie pasza do sida
'
' ==================================================================================================


' ver 0.4 2016-04-11
' ==================================================================================================
'
' zaciaganie recv na pelnej - wyglada to dobrze ale dane rozciagaja sie w dol
'
' ==================================================================================================


' ver 0.5 2016-04-12
' ==================================================================================================
'
' rozwazania nad sciaganiem z wizarda rqms'ow z odpowiednich pol
' pierwszej kolejnosci mysle nad zaciaganiem calosci, a moze jeszcze inaczej rozpoznawaniem danych
' wsadowych na podstawie zawartosci formuly
' w kolumnie Total_QTY czyli 16 kolumna (cos przed Q, to znaczy P)

'
'
' ==================================================================================================

' ver 0.6 2016-04-12
' ==================================================================================================
'
' proby pod cbala
'
'
' ==================================================================================================


' ver 0.7 2016-04-13
' ==================================================================================================
'
' proby pod cbala
'
'
' ==================================================================================================


' ver 0.8 2016-04-13
' ==================================================================================================
'
' praca nad layoutem coord list
' cos zle zaciaga dictionary kolejne
' zamiast key jako pn jest od razu key jako pus name - to be fixed
'
'
' ==================================================================================================

' ver 0.9 14-04
' ==================================================================================================
'
' layout dla coord list gotowy
' 0.91 dodanie kolumny RESP w CBALu
'
'
' ==================================================================================================

' ver 0.92 14-04
' ==================================================================================================
'
' zmiana zapisu logow
' recv status na text
'
'
' ==================================================================================================

' ver 1.00 20-04
' ==================================================================================================
'
' funkcjonalnosc coverage wstepnie przygotowana
'
'
' ==================================================================================================

' ver 1.01 21-04
' ==================================================================================================
'
' nowe kolory palety baku
'
'
' ==================================================================================================


' ver 1.02 21-04
' ==================================================================================================
'
' fix na kolorowaniu asnow / pusow
'
'
' ==================================================================================================

' ver 1.03 dev 06-06
' ==================================================================================================
'
' dodatkowe fixy na cov - dodatkowe kolumny
' plus moze jeszcze uproszczenie formularza.
' dodanie kolumn CBAL I POTENCIAL RECV
'
' blad na generowaniu coverage - fixed
' okazalo sie ze na jednym z obiektow nie wyczysciclem set o = Nothing
' normlanie nie musialbym tego robic jednak z powodu wstawienia on error
' gdy nie bylo danych ignorowal przypisanie powodujac maly bigos
'
' praca nad pus matchem w pierwszej fazie nie poszla za dobrze z powodu podwojnej petli do
' przez co makro zwieszalo sie podczas swojej pracy
' kod i tak nie zostal zapisany, zatem moge zaczac jeszcze raz
' przy okazji na kibelku postanowilem ze wroce rowniez do pierwotnej implementacji czystego arkusza PUSes
' i juz tam rozbije RECV TBD
'
'
' ==================================================================================================


' ver 1.04 dev 06-08
' ==================================================================================================
'
' odeszlem od pierwotnego zalozenia sprawdzenia recv tbd w szerszy sposob pozniej
' do razu zmodyfikowalem kod dla samego arkusza PUSes
' jedyne co zostalo do zaimplementowania to dodatnie pot recv dla arkusz PUS match
' same CBALE wsadzone sa juz poprawnie.
'
' powinny byc dwie petle a dalem jedna: jestem z siebie dumny (implementacja w PusHandler):
' Private Sub dodaj_do_siebie_potential_recvs(o As Worksheet)
'
' ==================================================================================================


' ver 1.05 dev 2016-06-14
' ==================================================================================================
'
' kolejny temat z coveragem - jednak fajnie  by bylo zeby potrafil zrobic jakas interakcje z zaciagnietym CBALem
' trzeba by ino rozpatrzec jakies rozsadne przeliczenie i skupic sie na current CW.
'
' aktualna opcja polega na patrzeniu na cala faze
' czy moze skrocic obraz do faktycznego ukladu coverage z pierwszym zrzutem dla current CW.
'
' ' dobra dobra - powyzej byly rozwazania, czyny ponizej:
' pierwsza opcja to podkreslanie na inny kolor w current week ebali ktore nie jest zbyt rowne z cbalem sciagnietym
'
' ==================================================================================================


' ver 1.06 dev 2016-06-27
' ==================================================================================================
'
' poprawiona logika definiowania pusow plus usuniecie pot recv
' proba testu logiki dla tylko dopasowanej nazwy.
' dodatkowo nie wzialem pod uwage ze recv rowniez moze miec wplyw na zmiene definicji linii
'
' ==================================================================================================


' -----------------------------------------------------------------------------------------------------------------
