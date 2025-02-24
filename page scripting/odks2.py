# -*- coding: utf-8 -*-
"""
Skrypt generuje nowy plik YAML o nazwie "odks2_recording.yml",
w którym sekcja "steps" zawiera 100 cykli kroków odpowiadających
jednemu wykonaniu operacji (odksięgowania dokumentu).

Przykładowy cykl (szablon):
  1. Invoke operacji ReverseTransaction na stronie "General Ledger Entries"
  2. Oczekiwanie, aż pojawi się strona "Reverse Transaction Entries"
  3. Invoke operacji Reverse na tej stronie
  4. Pojawia się strona Potwierdź, następnie invoke Yes,
  5. Zamknięcie stron i komunikat końcowy.

Wartości parametrów runtimeRef oraz runtimeId są generowane dynamicznie.
"""

# Początkowa część YAML – nagłówek
initial_content = """\
name: odks2 Recording
description: Test recording for odks2
start:
  profile: BUSINESS MANAGER
steps:
"""

# Szablon jednego cyklu kroków
cycle_template = """\
  - type: invoke
    target:
      - page: General Ledger Entries
        runtimeRef: {ledger_ref}
      - action: ReverseTransaction
    parameters: {{}}
    description: Invoke <caption>Wycofaj transakcję...</caption>
  - type: page-shown
    source:
      page: Reverse Transaction Entries
    modal: true
    runtimeId: {rev_ref}
    description: Page <caption>Wycofywanie zapisów transakcji - Konto K/G 139-09
      Transakcje karta-sprzedaż</caption> was shown.
  - type: invoke
    target:
      - page: Reverse Transaction Entries
        runtimeRef: {rev_ref}
      - action: Reverse
    parameters: {{}}
    description: Invoke <caption>Wycofaj</caption>
  - type: page-shown
    source:
      page: null
      automationId: 8da61efd-0002-0003-0507-0b0d1113171d
      caption: Potwierdź
    modal: true
    runtimeId: {conf_ref}
    description: Page <caption>Potwierdź</caption> was shown.
  - type: invoke
    target:
      - page: null
        automationId: 8da61efd-0002-0003-0507-0b0d1113171d
        caption: Potwierdź
        runtimeRef: {conf_ref}
    invokeType: Yes
    description: Invoke Yes on <caption>Potwierdź</caption>
  - type: page-closed
    source:
      page: null
    runtimeId: {conf_ref}
    description: Page <caption>Potwierdź</caption> was closed.
  - type: page-closed
    source:
      page: Reverse Transaction Entries
    runtimeId: {rev_ref}
    description: Page <caption>Wycofywanie zapisów transakcji - Konto K/G 139-09
      Transakcje karta-sprzedaż</caption> was closed.
  - type: message
    automationId: 8da61efd-0002-0003-0507-0b0d1113171d
    text: Zapisy zostały pomyślnie wycofane.
    description: "Message: <value>Zapisy zostały pomyślnie wycofane.</value>"
"""

# Budujemy treść końcową – zaczynamy od nagłówka
yaml_content = initial_content

# W pętli generujemy 100 cykli kroków.
# (Dla przykładu zmieniamy numery referencyjne – możesz je zmodyfikować według potrzeb)
for i in range(100):
    # Przykładowe generowanie referencji – dostosuj według własnych zasad!
    ledger_ref = f"b{390 + i}m"    # np. b390m, b391m, ...
    rev_ref    = f"b3is{i:02d}"      # np. b3is00, b3is01, ...
    conf_ref   = f"b3nr{i:02d}"      # np. b3nr00, b3nr01, ...
    
    yaml_content += cycle_template.format(
        ledger_ref=ledger_ref,
        rev_ref=rev_ref,
        conf_ref=conf_ref
    )

# Zapisujemy wynikowy YAML do pliku
output_file = "odks2_recording.yml"
with open(output_file, "w", encoding="utf-8") as f:
    f.write(yaml_content)

print(f"Plik '{output_file}' został utworzony z 100 cyklami kroków.")
