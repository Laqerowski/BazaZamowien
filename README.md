# BazaZamowien

## Opis projektu
Plik Excel zawiera stworzoną przeze mnie **Bazę Zamówień** produktów, która ma pomóc wprowadzać dane do arkusza. Projekt wykorzystuje makra VBA, aby zautomatyzować procesy związane z obsługą zamówień i zapewnić większą efektywność.

## Struktura projektu
W folderze `VBA` znajdują się dwa podfoldery:
- **Modules**: Zawiera moduły kodu VBA odpowiadające konkretnym funkcjom, które odpowiadają za modelowanie.
- **Objects**: Zawiera kod VBA powiązany z arkuszami (`Arkusz1`, `Arkusz2`) i odpowiadający za wywołania konkretnych modułów w reakcji na zmiany w danych (`Worksheet_Change`).

## Jak używać
1. Otwórz plik Excel i upewnij się, że obsługa makr jest włączona.
2. Importuj moduły i obiekty arkuszy:
   - W edytorze VBA (`Alt + F11`), wybierz **File → Import File...** i załaduj pliki z folderów `Modules` i `Objects`.
3. Wprowadź dane zamówień do odpowiednich arkuszy Excela.
4. Makra uruchomią się automatycznie przy wprowadzaniu lub zmianie danych, przekształcając konkretne komórki.
