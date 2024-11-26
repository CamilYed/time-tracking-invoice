# Automatyczny Generatory Raportów

Ten projekt to prosty generator raportów na podstawie danych osobowych oraz godzin pracy w danym miesiącu. Skrypt pobiera informacje o polskich świętach z publicznego API, tworzy tabelę z godzinami pracy i automatycznie oblicza sumę godzin oraz kwotę na podstawie stawki godzinowej.

---
## Wymagania

Skrypt wymaga zainstalowania następujących bibliotek Python:

- `requests` — do pobierania danych z API o świętach.
- `click` — do obsługi interfejsu wiersza poleceń.
- `openpyxl` — do generowania plików Excel.

Skrypt samodzielnie instaluje brakujące pakiety, jeśli nie zostały wcześniej zainstalowane.

---

## Jak używać

### 1. Konfiguracja danych osobowych

Aby skonfigurować dane osobowe, uruchom następujące polecenie:

```bash
python main.py configure
```
Podaj wymagane dane:
- Imię i nazwisko
- Rola/Stanowisko
- Numer zamówienia/umowy T&M
- Stawka godzinowa (PLN)

Te dane zostaną zapisane w pliku person_config.json.

---

### 2. Generowanie raportu
Aby wygenerować raport na podstawie danych za wybrany rok i miesiąc, uruchom następujące polecenie:

```bash
python generate-hours-summary.py generate --year 2024 --month 12
```

Jak działa skrypt:
Pobieranie świąt: Skrypt automatycznie pobiera dane o świętach w Polsce z API: https://date.nager.at.
Tworzenie raportu: Na podstawie danych o świętach i godzinach pracy generuje raport w Excelu.
Obliczenia: Oblicza sumę godzin pracy i kwotę na podstawie stawki godzinowej.
Skrypt automatycznie wypełnia dni tygodnia oraz oznacza weekendy i święta. Zawiera również podsumowanie godzin oraz kwoty na końcu.
