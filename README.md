# Console Spreadsheet (Mini Excel) — C#

A **C# console application** that simulates a small spreadsheet (mini Excel) with basic cell operations such as assigning values, clearing, inserting rows/columns, copy/cut, arithmetic operations, and simple encryption.

The sheet is displayed in the console with **column letters (A, B, C...)** and **row numbers (1, 2, 3...)**, similar to Excel.

---

## Key Concepts

- The spreadsheet stores:
  - `cells[r, c]` → the **value** (string)
  - `types[r, c]` → the **type** (`integer`, `string`, `unassigned`)
- Default visible size:
  - **8 rows × 5 columns**
- Maximum supported storage size:
  - Up to **15 rows × 10 columns** (arrays are fixed size)

---

## Features

### Cell Management
- **AssignValue**: set a cell value + type (`integer` / `string`)
- **View Cell**: display full content of a cell
- **ClearCell**: clear one cell
- **ClearAll**: reset the whole sheet

### Structure Editing
- **AddRow**: insert a row `up/down` from a given row index
- **AddColumn**: insert a column `left/right` from a given column letter

### Clipboard-Like Tools
- **Copy**: copy a single cell to another
- **CopyRow / CopyColumn**
- **CutCell / CutRow / CutColumn**: move content and clear source

### Operations
- **Add**
  - If all inputs are integers → numeric sum (2 or 3 cells)
  - If any input is string → concatenation + case mode (`up` / `low`)
- **Subtract**
  - int - int → numeric subtraction
  - string - string → remove substring occurrences
  - string & int → remove a specific ASCII character (integer must be **[33..126]**)
- **Multiply**
  - int * int → numeric multiplication
  - string * int → repeat string
    - negative integer → repeat the **reversed string**
  - string * string → **not allowed**
- **Divide**
  - int / int → integer division (division by 0 not allowed)
  - string / int → takes part of the string:
    - positive → from the beginning
    - negative → from the end
  - division by 0 not allowed

### Encrypt
- Caesar-style shift on a string using an integer shift value
- Works only when one cell is `string` and the other is `integer`
- Integer shift range must be **[-20..30]**

### Save to File
- On exit (`0`), the program saves the current sheet view to:
  - `spreadsheet.txt`

---

## Menu

1. AssignValue  
2. ClearCell  
3. ClearAll  
4. AddRow  
5. AddColumn  
6. Copy  
7. CopyColumn  
8. CopyRow  
9. CutCell  
10. CutColumn  
11. CutRow  
12. Multiply  
13. Add  
14. Divide  
15. Subtract  
16. Encrypt  
17. View Cell  
0. Exit (auto-saves)

---

## How to Run

### .NET CLI
```bash
git clone https://github.com/<your-username>/<your-repo>.git
cd <your-repo>
dotnet run
