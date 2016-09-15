Attribute VB_Name = "VersionModule"

' 2016-09-09 1.00
' ===========================================================================================
' pierwsze udane testy jesli chodzi o generowanie maili jednak ich ilosc jest przytlaczajaca
' i dobrze by bylo sprawdzic czy ilosci jakie zostale wygenerowane sa zgodne z tym co jest w
' wyparowanym Wizardzie
' ===========================================================================================
' 2016-09-09 0.99
' ===========================================================================================
' 19 milestone
' yellow box + fixy na material order sheet - thin lines + usuniecie tabeli z htmlbody
' starego jesli nic nie zostalo wsadzone
' ===========================================================================================
' 2016-09-09 0.98
' ===========================================================================================
' 18 milestone
' rozszerzenie kompetencji narzedzia na generowanie maili dopasowanych do typu proj:
' Major / MY
' PSA , BIW itd.
' przez co zostal rozszerzony formularz wyboru i co za tym idzie maile moge byc bardziej
' dostosowane do ukladu
' narazie obrazek do maila mnie pokonal (wiec schowalem arkusz z nim)
' ale jeszcze do tego wroce
' ===========================================================================================

' 2016-08-30 0.97
' ===========================================================================================
' 17 milestone
' - dodanie mozliwosci generowania samej tabeli
' - pojedynczy generowanie tresci - po co za kazdym razem to samo
' ===========================================================================================

' 2016-08-30 0.96
' ===========================================================================================
' 16 milestone - work only with visible data
' ===========================================================================================

' 2016-08-29 0.95
' ===========================================================================================
' proba przyspieszenia parsera text 2 html
' ===========================================================================================

' 2016-08-26 0.94
' ===========================================================================================
' skrocenie czasu pracy przez wyrzucenie potrzebnie podzielonych wordow
' ===========================================================================================

' 2016-08-08
' BUCKET 0.91 prototype
' ===========================================================================================

' 2016-08-24 0.93
' --------------------------
' poprawiony material order

' 2016-08-08
' BUCKET 0.91 prototype
' ===========================================================================================
'
'
'
' po meetingu
' w drafcie maila - nie jest potrzebny PLT ani PROJ name
' part odrder conf nie jest juz potrzebny
'
'
' ===========================================================================================

' 2016-08-08
' BUCKET 0.9 prototype
' ===========================================================================================
'
'
' Order w postaci order.htm jest nieedytowalny, co nie jest zbyt wygodne.
' Wersja 0.9 wykorzystuje dodatkowy arkussz material order pod wpisanie w postaci excelowej
' tabeli zamowien
' ktora bedzie dostepna zarowno do 10 linii bezposrednio w mailu, ale i pliku zalaczonym do maila.
'
'
' ===========================================================================================



' 2016-08-04
' BUCKET 0.8 prototype
'
' zmiana koncepcji szablonu maila
' wszystko w zgodzie z nowymi zalozeniami:
' zostawiamy jako tako zalaczniki
' zmieniamy nieco glowne punkty (arkusz info) przestajemy wpisywac tresc bezposrednio w komorki
' stworzylem na potrzeby ole objecty worda,
' w ktore wstawione tresci z szablonowego mail od piotera (embed objects)

' tabela zamowien musi dynamicznie miec szanse zmieniac sie w plik excelowy jesli
' sie okaze ze jest wiecej niz 10 orderow - to jeszcze do zweryfikowania
' ===========================================================================================
' 2016-08-01, 2016-08-02, 2016-08-03
' BUCKET 0.7 prototype
'
' zmiana koncepcji szablonu maila
' wszystko w zgodzie z nowymi zalozeniami:
' zostawiamy jako tako zalaczniki
' zmieniamy nieco glowne punkty (arkusz info) przestajemy wpisywac tresc bezposrednio w komorki
' stworzylem na potrzeby ole objecty worda,
' w ktore wstawione tresci z szablonowego mail od piotera (embed objects)

' tabela zamowien musi dynamicznie miec szanse zmieniac sie w plik excelowy jesli
' sie okaze ze jest wiecej niz 10 orderow - to jeszcze do zweryfikowania
' ===========================================================================================

' 2016-03-08
' BUCKET 0.6 prototype
' jeszcze go nie zaczalem - jednak plan jest taki aby
' zmienic uklad dodawania info
' sa 4 podpunkty zaczyanajac od GENERAL INFO
' PAYMENTS
' mysle zeby nie parsowac wybitnie takich info tylko dodatkowy plik ogarnac w postaci maila polozonego ot tak
'
' dodatkowo kolejny meeting z fma w koncu jakie zalaczniki zostana dodane do koncowego maila.
' ===========================================================================================


' 2016-02-17
' BUCKET 0.5
' dodanie zalacznika nowego MATERIAL ORDER
' - dostawca nie chce samego maila chce dostac zalacznik
' ===========================================================================================

' 2016-02-04
' BUCKET 0.4
' proba usuniecia wszystkich warningow
' odchudzenie paddingow
' oraz roszerzenie szerokosci tabeli zamowien do 800px
' ===========================================================================================

' 2016-02-03
' BUCKET 0.3
' attaching files
' ===========================================================================================

' 2016-02-03
' BUCKET 0.2
' ustawienie nowych arkuszy wraz z danymi, ktore potem beda wydzielone jako osobne pliki
' i zalaczone do maili
' plus pliki side'owo wsadzoene
' ===========================================================================================

' 2016-02-02
' BUCKET 0.1
' pierwszy prototyp z surowymi mailami bez konkretow oprocz wybierania danych z wizarda
' ===========================================================================================
