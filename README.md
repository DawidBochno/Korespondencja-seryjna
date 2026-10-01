# Korespondencja seryjna (DOCX + XLSX)

Program **lokalny** — działa w całości na Twoim komputerze i nigdzie nie
wysyła danych. Z jednego szablonu pisma w Wordzie i listy adresatów
w Excelu tworzy osobne pismo dla każdego adresata: zawiadomienia, decyzje,
zaświadczenia, wezwania, podziękowania.

## Jak to działa

1. **Szablon DOCX.** Zwykły dokument Worda, w którym zmienne miejsca
   oznaczasz nazwą pola w klamrach: `{imie}`, `{nazwisko}`, `{adres}`,
   `{kwota}`.
2. **Dane XLSX.** Pierwszy wiersz pierwszego arkusza zawiera nazwy pól
   (bez klamer), a każdy kolejny wiersz to jedno pismo. Puste wiersze są
   pomijane.
3. Program wstawia wartości w miejsce pól i zapisuje osobny `.docx` dla
   każdego wiersza, a jeśli zaznaczysz tę opcję, także `.pdf`.

Przykład jest w folderze [`przyklad/`](przyklad/): `szablon.docx`
(zawiadomienie o opłacie) i `dane.xlsx` (3 fikcyjnych adresatów). Program
otwiera się z tymi plikami wpisanymi w pola, więc wystarczy kliknąć
**Generuj pisma**.

| szablon | arkusz | wynik |
|---------|--------|-------|
| `Szanowna Pani {imie} {nazwisko}` | `imie: Anna`, `nazwisko: Nowak` | `Szanowna Pani Anna Nowak` |

### Zasady

- **Formatowanie zostaje** (czcionka, pogrubienie, tabele). Wartość
  przejmuje wygląd miejsca, w którym stało pole. Pole może stać w treści,
  tabeli, nagłówku, stopce i polu tekstowym.
- **Wielkość liter w nazwach pól nie ma znaczenia**: `{Imie}` = `{imie}` =
  kolumna `IMIE`.
- **Adres w kilku liniach**: w komórce Excela przejdź do nowej linii
  (Alt+Enter), a w piśmie pojawi się złamanie wiersza.
- **Daty** z Excela są wpisywane jako `15.10.2026`.
- **Liczby** są wpisywane zgodnie z formatem komórki w Excelu. Format
  `0,00` daje `42,50`, format z separatorem tysięcy daje `1 234,50`.
  Przy formacie „Ogólny” program wpisze `42,5`.
- **`{nr}`** to numer wiersza (1, 2, 3…). Działa nawet bez takiej kolumny.
- Pola z szablonu, których nie ma w arkuszu, zostają w piśmie bez zmian,
  a w logu pojawia się ostrzeżenie. Program sprawdza to przed generowaniem.

### Nazwy plików

Pole **Nazwa pliku** przyjmuje te same `{pola}`, np. `{nr}_{nazwisko}_{imie}`
daje `1_Nowak_Anna.docx`. Znaki niedozwolone w nazwach plików są
zamieniane na `_`. Jeśli nazwy się powtarzają (np. to samo nazwisko),
kolejne pliki dostają końcówki `_2`, `_3`.

### PDF

Opcja **Utwórz też PDF** używa zainstalowanego **Microsoft Word**, więc PDF
wygląda dokładnie jak wydruk z Worda. Word jest uruchamiany raz na całą
serię, niewidocznie. Żeby dostać jeden plik do wydruku całej serii, połącz
PDF-y programem *Narzędzia PDF* (operacja **Połącz**).

## Szybki start

1. `install.bat` — instaluje biblioteki (`python-docx`, `openpyxl`,
   `pywin32`) i uruchamia self-test.
2. `uruchom.bat` → wskaż szablon i arkusz → **Generuj pisma**.
3. Wyniki są w folderze `OUTPUT`.

Wymaga Pythona 3.9+ z opcjami „Add python.exe to PATH” i „tcl/tk and IDLE”.

## Ograniczenia

- Obsługiwany jest tylko pierwszy arkusz pliku XLSX. Format `.xls` trzeba
  najpierw zapisać jako `.xlsx`.
- Nie da się tworzyć warunków (np. „Pan/Pani” zależnie od płci). Dodaj
  w arkuszu kolumnę `{zwrot}` z gotowym tekstem.
- Opcja PDF wymaga Microsoft Word. Bez niego powstają same pliki DOCX.

## Testy

```bash
python korespondencja.py --selftest         # logika (bez Worda)
python korespondencja.py --selftest --pdf   # dodatkowo konwersja PDF przez Worda
```

Test sprawdza pole rozcięte przez Worda na kilka fragmentów formatowania,
pola w nagłówku i tabeli, adres wieloliniowy, daty, formaty liczb, pomijanie
pustych wierszy, nieznane pola i powtarzające się nazwy plików.
