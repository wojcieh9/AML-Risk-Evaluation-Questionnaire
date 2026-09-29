# Weryfikacja eksportera CBI v1.0

Data: 28 września 2026.

## Wykonane kontrole

| Kontrola | Wynik |
|---|---|
| Pokrycie identyfikatorów pojedynczych odpowiedzi z XSD | 456/456, bez duplikatów |
| Zgodność typów/atrybutów pojedynczych odpowiedzi z Individual Question Instructions | 456/456 |
| Pokrycie tabel | 9/9, wszystkie atrybuty XSD |
| Kolejność sekcji | Wszystkie 12 sekcji według sekwencji XSD |
| Adresy komórek i etykiety kontrolne | Sprawdzone w formularzu bazowym i obu przykładach XLSX |
| Przykład Completed | Odtworzono 530 rekordów XML; porównano wszystkie atrybuty z przykładowym XML banku |
| Przykład Completed Blank | Odtworzono 520 rekordów XML; porównano wszystkie atrybuty z przykładowym XML banku |
| Walidacja testowego XML po jawnych poprawkach danych testowych | Oba pliki poprawne w lxml XMLSchema i Windows MSXML 6.0 |
| Testy negatywne w MSXML 6.0 | Odrzucenie błędnej wielkości liter w słowniku, przekroczenia zakresu integer, nieistniejącej daty, procentu >1 i nadmiernej precyzji |
| Diagnostyka węzła źródłowego | MSXML validateNode wykrywa błąd w konkretnym rekordzie |
| Funkcje pomocnicze z kodu modułu | 10 testów normalizacji liczb, wyrażeń regularnych i kodowania XML w zgodnym podzbiorze VBScript |
| Niezmieniony XSD | Kopia identyczna bajtowo ze źródłem |

Porównanie atrybutów ujawniło różnice między samymi przykładami XLSX i XML banku oraz równoważne zapisy liczb, np. `0` i `0.0`. Nie użyto błędnych wartości XML jako wzorca poprawnych odpowiedzi.

## Jawne poprawki wyłącznie w syntetycznych danych testowych

- Completed: LEI `LAG` i `LAO` zastąpiono 20 zerami; `LIX` ustawiono na `N/A`, `LIY` na `No_transaction_monitoring`.
- Completed Blank: `LIX` ustawiono na `N/A`, `LIY` na `No_transaction_monitoring`; `LGG`, `LGH`, `LPL` ustawiono na `No`. Wartości LEI `LAG`, `LAO` ustawiono na 20 zer.
- Nie zmieniono żadnego źródłowego pliku banku. Makro nie stosuje tych poprawek do danych użytkownika — zgłasza błędy do rozstrzygnięcia przez właściciela danych.

Wartości wybrane do testu nie są rekomendacją odpowiedzi merytorycznej ani gotowym sprawozdaniem.

## Czego te testy nie potwierdzają

Nie uruchomiono ani nie skompilowano VBA w Excelu, ponieważ uruchomienie Excela w tle nie zostało zatwierdzone. Testy struktury XML wykonał niezależny kod odczytujący tę samą mapę komórek; nie jest to test końcowy makra. Testy podzbioru VBScript nie zastępują kompilacji VBA.

Pierwsza instalacja wymaga `Debug → Compile VBAProject` oraz uruchomienia `Sprawdz_Dane_CBI`. Przed użyciem produkcyjnym należy wykonać eksport na kontrolowanym komplecie danych i porównać odpowiedzi w XML z arkuszem.

Nie wykonano wysyłki do Portalu CBI ani pełnej weryfikacji reguł merytorycznych i uzgodnień między sekcjami.
