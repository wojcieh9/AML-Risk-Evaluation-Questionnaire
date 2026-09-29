# Makro eksportu XML — CBI Fund Administrators v1.0

Pakiet zawiera moduł VBA do istniejącego formularza banku. Nie jest to gotowy plik XLSM.
Wymagany jest desktopowy Microsoft Excel na Windows z dopuszczonymi makrami.
Makro korzysta z systemowych MSXML 6.0 i ADODB; nie wymaga Pythona, internetu ani dodatkowych referencji w edytorze VBA.

## Instalacja — jednorazowo

1. Rozpakuj cały ZIP do lokalnego folderu.
2. Otwórz **kopię** pliku `Funds Administrators - Risk Evaluation Questionnaire v1.0.xlsx` z pakietu banku.
3. Wybierz **Plik → Zapisz jako → Skoroszyt programu Excel z obsługą makr (*.xlsm)**. Zapisz go w folderze zawierającym `Fund Admin - XSD v1.0.xsd`.
4. Naciśnij **Alt+F11**. W panelu projektu zaznacz swój skoroszyt XLSM.
5. Wybierz **File / Plik → Import File / Importuj plik** i wskaż `CBI_REQ_Export.bas`. Powstanie moduł `CBI_REQ_Export`. Nie wklejaj pliku do modułu arkusza ani `ThisWorkbook`.
6. W edytorze wybierz **Debug → Compile VBAProject**. Zapisz skoroszyt i wróć do Excela przez Alt+F11.
7. Uruchom przez **Alt+F8 → Sprawdz_Dane_CBI → Uruchom**. Niewypełniony formularz powinien zgłosić brakujące odpowiedzi i utworzyć raport TXT.

Import ręczny nie wymaga włączania opcji „Ufaj dostępowi do modelu obiektowego projektu VBA”. Przy blokadzie makr przez politykę firmową zastosuj firmowy proces zatwierdzania lub podpisania kodu; nie zmieniaj globalnie zabezpieczeń.

Opcjonalny przycisk: **Deweloper → Wstaw → Przycisk (formant formularza)**, przypisz `Eksportuj_XML_CBI`. Umieść go w wolnym miejscu bez przesuwania istniejących komórek.

## Uzupełnianie formularza

- W komórkach pojedynczych odpowiedzi zastąp kody, np. `LAA`, `LBV`, właściwą odpowiedzią. Makro zna ich oryginalne adresy.
- W tabelach pozostaw wiersze nagłówków z kodami pól i wpisuj dane poniżej. Nie zmieniaj nazw arkuszy, układu kolumn, etykiet ani położenia pojedynczych pytań.
- Możesz dopisywać rekordy pod tabelami na arkuszach z danymi funduszy, krajów i monitorowania. Eksporter odczytuje wszystkie wypełnione wiersze, również poza początkowo sformatowanym obszarem. Limity XSD nadal obowiązują: do 250 rekordów dla tabel zmiennych z wyjątkiem monitorowania transakcji.
- Tabela obecności międzynarodowej obejmuje 30 krajów EOG, a tabela sektorów — 27 sektorów. Wypełnij wszystkie ich wiersze.
- Całkowicie puste wiersze w tabelach zmiennych są pomijane. Częściowo wypełnione wiersze są błędem. Pusta tabela również jest błędem; jeśli sekcja nie dotyczy, wpisz wymagany przez instrukcję rekord „nie dotyczy”.
- Pustka nie oznacza zera. Makro nie dopisuje odpowiedzi, zer ani „N/A”.
- Odpowiedzi słownikowe muszą dokładnie odpowiadać XSD, łącznie z wielkością liter: `Yes`, nie `YES` ani `yES`.
- Kody krajów `00` oraz LEI przechowuj jako **tekst**. LEI „nie dotyczy” to 20 zer: `00000000000000000000`; nie `N/A` i nie liczba zero. Format komórki ustaw przed wpisaniem wartości lub poprzedź ją apostrofem.
- Daty: rzeczywista data Excela bez godziny albo tekst `YYYY-MM-DD`. Makro obsługuje oba systemy dat Excela: 1900 i 1904.
- Procenty: liczba od 0 do 1, np. `0,047` albo komórka z wartością `4,7%` przechowywaną jako 0,047. Tekst `4,7%` nie jest obsługiwany. Maksymalnie 6 miejsc po przecinku.
- Liczby zapisane jako tekst: bez spacji i separatorów tysięcy; dopuszczalny przecinek lub kropka dziesiętna. Makro nie zaokrągla nadmiernej precyzji. Długie liczby mogą wymagać zapisania jako tekst, ponieważ Excel zachowuje jedynie około 15 cyfr znaczących.
- Teksty mają limity instrukcji: 100, 500 lub 2000 znaków. Są one celowo ostrzejsze niż margines dodatkowych 50 znaków w XSD. Walidacja XSD dodatkowo sprawdza ograniczenia znaków.

**Uwaga na dane przykładowe:** także podstawowy formularz banku zawiera przykłady w tabelach, m.in. GB/FR oraz nazwy reguł monitorowania. Zastąp je własnymi danymi lub usuń. Poprawnie sformatowany przykład może przejść kontrolę techniczną — makro nie potrafi rozpoznać, czy dane są prawdziwe.

## Każdy eksport

1. Uzupełnij odpowiedzi, zapisz skoroszyt i uruchom `Sprawdz_Dane_CBI`.
2. Przy błędach otwórz `CBI_validation_YYYYMMDD_hhmmss.txt` w folderze skoroszytu. Raport wskazuje REF i komórkę albo wiersz tabeli. Popraw dane i ponów kontrolę.
3. Uruchom `Eksportuj_XML_CBI`. Podaj kod instytucji używany w Portalu i datę raportową, np. `2026-12-31`, a nie datę wysyłki.
4. Wybierz folder zapisu. Zachowaj wygenerowaną nazwę `KOD_YYYYMMDD_A12.xml`.
5. Po sprawdzeniu merytorycznym prześlij XML przez Portal CBI i wykonaj `Finalise` oraz `Sign-Off`.

Eksport powstaje dopiero po przejściu kontroli wejścia i walidacji całego dokumentu względem XSD. Przy błędach nowy XML nie jest zapisywany. **Istniejący plik z poprzedniego eksportu pozostaje niezmieniony** — nie traktuj go jako aktualnego po nieudanej próbie.

Jeżeli XSD nie leży obok skoroszytu, makro poprosi o jego wskazanie. Nie zmieniaj schematu banku. Nowa wersja formularza lub XSD wymaga ponownego sprawdzenia mapowania i aktualizacji eksportera.

## Zakres i ograniczenia

- Obsługiwany jest układ bankowego formularza Fund Administrators v1.0: 456 odpowiedzi pojedynczych, 9 tabel i wszystkie 12 sekcji XML.
- Makro sprawdza etykiety przy pytaniach i nagłówki tabel, brakujące odpowiedzi, powtórzone klucze wierszy, limity rekordów, typy, słowniki, zakresy, daty i precyzję. Generuje UTF-8, właściwą przestrzeń nazw i pełne znaczniki zamykające zamiast `/>`.
- Odpowiedzi są odczytywane z tego skoroszytu, do którego zaimportowano moduł (`ThisWorkbook`). Nie importuj go do PERSONAL.XLSB ani osobnego pustego skoroszytu.
- Źródłowe odpowiedzi nie są modyfikowane. Makro przelicza arkusze przed odczytem. Działa lokalnie; nie przesyła danych i nie loguje się do Portalu.
- Walidacja obejmuje poprawność techniczną i wskazane kontrole kompletności, **nie wszystkie zależności merytoryczne**, np. uzgodnienia sum między sekcjami, zakres podmiotów grupy, zgodność ze źródłami i interpretację „nie dotyczy”. Nie gwarantuje przyjęcia przez dodatkowe kontrole Portalu.
- Pakiet przetestowano pod kątem kompletności mapowania i zgodności wynikowej struktury z XSD w lxml oraz MSXML 6.0. **Nie uruchomiono samego VBA w Excelu** — uruchomienie Excela w tle nie zostało zatwierdzone. Po instalacji konieczne są kompilacja i pierwszy test w docelowym Excelu.

## Niespójności przykładów banku

Nie poprawiamy ich automatycznie w danych użytkownika. Raport walidacji wymaga świadomej poprawy odpowiedzi:

- `LIX`: w przykładach występuje `No_transaction_monitoring`, choć pole przyjmuje `Yes`, `No`, `N/A`.
- `LIY`: przykład arkuszowy używa `N/A`; XSD dla tego pola przewiduje m.in. `No_transaction_monitoring`, ale nie `N/A`.
- Przykładowy wypełniony XLSX zawiera `N/A` dla LEI `LAG` i `LAO`; wymagany LEI ma 20 znaków.
- Przykładowy „Completed Blank” XLSX ma `N/A` w `LGG` i `LGH` oraz datę w `LPL`, mimo że są to odpowiedzi `Yes`/`No`.
- Przykładowy wypełniony XML ma `yES` w `LBV` i `0.0` w dwóch polach całkowitych. Makro odczytuje Excel i nie kopiuje tych wartości z XML-a.

## Pliki w pakiecie

- `CBI_REQ_Export.bas` — moduł do importu.
- `Fund Admin - XSD v1.0.xsd` — niezmieniony schemat banku.
- `mapowanie_pol.csv` — audytowalne powiązanie pojedynczych odpowiedzi z komórkami i atrybutami XML; nie jest potrzebne do uruchomienia.
- `mapowanie_tabel.csv` — zakresy i kody tabel; nie jest potrzebne do uruchomienia.
- `WYNIKI_TESTOW.md` — zakres wykonanej weryfikacji.
- `INSTRUKCJA.md` — ten dokument.

Materiały źródłowe: https://www.centralbank.ie/regulation/anti-money-laundering-and-countering-the-financing-of-terrorism/sector-specific-ml-tf-risk-evaluation-questionnaire
