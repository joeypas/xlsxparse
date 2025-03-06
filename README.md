# xlsxparse

xlsxparse is a Python command-line application that extracts cell references from an Excel file and outputs their locations to a JSON file. The tool provides two main commands:

- `create`: Parses an Excel (`.xlsx`) file for references and exports them to a JSON file.
- `search`: Searches the generated JSON file for specific references.

## Installation

You can install xlsxparse via pip:

```sh
pip install xlsxparse-joeypas
```

Ensure you have Python 3.8+ installed.

## Usage

### Create Command

Extracts references from an Excel file and saves them to a JSON file.

```sh
xlsxparse create <xlsx_file> [--sheet <sheet_name>] --output <json_file>
```

#### Parameters:
- `<xlsx_file>`: Path to the Excel file.
- `--sheet <sheet_name>` (optional): Name of the specific sheet to parse. If omitted, all sheets are parsed.
- `--output <json_file>`: (optional): Path to save the extracted references in JSON format.

#### Example:
```sh
xlsxparse create data.xlsx --sheet Sheet1 --output refs.json
```

### Search Command

Searches the JSON file for specific references.

```sh
xlsxparse search [OPTIONS] STRING [FILE]
```

#### Arguments:
- `STRING` (required): String to search for.
- `FILE` (optional): Path to the file. Defaults to `output.json`.

#### Options:
- `--stype [metric|sheet-metric|ref-file|ref-file-sheet]` (default: `metric`): Search type.
- `--help`: Show help message and exit.

#### Example:
```sh
xlsxparse search "Book2.xlsx, Sheet2" refs.json --stype ref-file-sheet
```

## Output Format

The output JSON file follows this structure:

```json
[
  {
    "Sheet": "Sheet1",
    "Metric": [
      "Test",
      "Three"
    ],
    "Cell": "B4",
    "Formula": "=B2+B3",
    "References": [
      {
        "sheet": "Sheet1",
        "cell": "B2"
      },
      {
        "sheet": "Sheet1",
        "cell": "B3"
      }
    ]
  },
  {
    "Sheet": "Sheet2",
    "Metric": [
      null,
      null
    ],
    "Cell": "A2",
    "Formula": "=SUM(30,[1]Sheet1!D12)",
    "References": [
      {
        "file": "Book2.xlsx",
        "sheet": "Sheet1",
        "cell": "D12"
      }
    ]
  }
]
```

## Dependencies

- Python 3.8+
- openpyxl

## Contributing

1. Fork the repository.
2. Create a new branch (`feature-branch`).
3. Commit your changes.
4. Push to your branch and open a Pull Request.

## License

This project is licensed under the MIT License.

