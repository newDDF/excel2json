# spreadsheet2json

A lightweight, utility-focused desktop tool designed to convert Excel (`.xlsx`, `.xls`) and LibreOffice Calc (`.ods`) files into formatted JSON files. Built with Python and Tkinter, it provides a simple graphical interface for Windows users and a standalone command-line tool for Linux servers.

## Features

- **Zero External Dependencies**: The parsing core relies solely on standard built-in Python XML/Zip libraries, minimizing binary payload footprint.
- **Dual Format Support**: Reads internal data structures from both Microsoft Excel and LibreOffice spreadsheet tables.
- **Collision Strategy Management**: Offers automated overwriting, conditional prompt notifications, or sequential indexing filenames when encountering pre-existing target file pathways.
- **Cross-Platform Deployable**: Optimized to bundle down into ultra-lean standalone execution binaries (~11MB for Windows GUI and ~8MB for Linux CLI).

## Installation & Setup

### Prerequisites

To execute the raw script environments directly, ensure you have Python 3.8+ installed along with the default graphic loop libraries (`tkinter`).

```bash
# Clone the repository
git clone https://github.com
cd excel2json
```

## Compilation

To bundle these scripts into standalone binaries that do not require an active Python interpreter installed on host deployment machines, run `pyinstaller`:

### Windows Binary (GUI)
```powershell
# Cleans workspace paths and compiles into a single background-window package
Remove-Item -Path .\build, .\dist, .\*.spec -Recurse -ErrorAction SilentlyContinue
pyinstaller --clean -y --onefile -w --icon="logo.ico" --add-data="logo.ico;." --name="Matrix_Converter" excel_to_json_cyber.py
```

### Linux Binary (CLI)
```bash
# Compiles the lightweight command line matrix workflow on local Linux host boxes
pyinstaller --onefile --name="Matrix_Converter_Linux" excel_to_json_linux.py
```

## Usage

### Windows (Graphic Interface)
Simply execute `Matrix_Converter.exe` via file explorers or launch environments, load your source file grid via the selector prompt, specify your duplication mitigation handling, and click `EXECUTE`.

### Linux (Terminal Console)
Grant operational binary runtime privileges and feed files down input argument arrays:
```bash
chmod +x ./Matrix_Converter_Linux
./Matrix_Converter_Linux file1.xlsx path/to/file2.ods --rename
```

## License

Distributed under the terms and compliance specifications of the [MIT License](LICENSE).
