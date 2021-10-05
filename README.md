# ExcelManager

Program do zarządzania danymi z excela. W obecnej wersji V1 zawarte są funkcje:

1) getExcelWorksheetNames - pobieranie nazw arkuszy z pliku excel
2) getExcelWorksheetColumnNames - pobieranie nazw kolumn z danego arkusza excel
3) getExcelWorksheetData - pobieranie danych z arkusza

Aby uruchomić program, należy wywołać go z wiersza poleceń (CMD) wraz z odpowiednimi parametrami.

1) Ścieżka pliku
2) Nazwa arkusza
3) Początek zakresu (np A1)
4) Koniec zakresu (np b2)
5) Rodzaj komendy (niewykorzystywane z obecnej wersji, odpowiedzialne za wybór odpowiedniej metody)
6) Wymagania służące do filtrowania. Powinny być podawane w ciągu znaków w konwencji: "nazwa_kolumny1"="wartość1","nazwa_kolumny2"="wartość2"

Przykładowe wywołanie: C:\ExcelManager.exe C:\lista.xlsx Arkusz1 A1 B5 wyszukajWiersz Rok=2015,Nazwa=ABC
