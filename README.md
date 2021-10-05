# ExcelManager

Program do zarządzania danymi z excela. W obecnej wersji V1 zawarte są funkcje:

1) getExcelWorksheetNames - pobieranie nazw arkuszy z pliku excel
2) getExcelWorksheetColumnNames - pobieranie nazw kolumn z danego arkusza excel
3) getExcelWorksheetData - pobieranie danych z arkusza

Aby uruchomić program, należy wywołać go z wiersza poleneć (CMD) wraz z odpowiednimi parametrami.

Parametr 1 - ścieżka pliku
Parametr 2 - nazwa arkusza
Parametr 3 - początek zakresu (np A1)
Parametr 4 - koniec zakresu (np b2)
Parametr 5 - rodzaj komendy (niewykorzystywane z obecnej wersji, odpowiedzialne za wybór odpowiedniej metody)
Parametr 6 - wymagania służące do filtrowania. Powinny być podawane w ciągu znaków w konwencji: "nazwa_kolumny1"="wartość1","nazwa_kolumny2"="wartość2"

Przykładowe wywołanie: C:\ExcelManager.exe C:\lista.xlsx Arkusz1 A1 B5 wyszukajWiersz Rok=2015,Nazwa=ABC
