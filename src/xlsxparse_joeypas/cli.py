import typer
import json
from xlsxparse_joeypas.parse import parse_all_sheets, parse_single_sheet
from xlsxparse_joeypas.search import search_cell, search_ref_file, search_ref_sheet, search_ref_file_sheet
from typing_extensions import Annotated
from typing import Optional
from enum import Enum

__version__ = "0.1.0"

def version_callback(value: bool):
    if value:
        print(f"xlsxparse version: {__version__}")
        raise typer.Exit()

app = typer.Typer(no_args_is_help=True)

@app.command()
def create(
    file: Annotated[str, typer.Argument(help="Path to the .xlsx WorkBook")],
    sheet_name: Annotated[str, typer.Argument(help="Name of the sheet to parse (if none provided, parse all sheets)")] = None,
    to_json: Annotated[bool, typer.Option("--json", help="Output file format as json.")] = False,
    verbose: Annotated[bool, typer.Option("--verbose", help="Write output to console")] = False,
    output_file: Annotated[str, typer.Option(help="Name of file to output")] = "output",
    version: Annotated[Optional[bool], typer.Option("--version", callback=version_callback)] = None,
):
    if to_json:
        with open(output_file + ".json", "w+") as txt_file:
            out = []
            if sheet_name:
                refs = parse_single_sheet(file, sheet_name)
                for cell, data in refs.items():
                    out.append({
                        'Sheet': sheet_name, 'Metric': data['names'],
                        'Cell': cell, 'Formula': data['formula'], 'References': data['references']
                    })
                    if verbose:
                        print(f"Cell: {cell}, Formula: {data['formula']}, Refrences: {data['references']}")
            else:
                refs = parse_all_sheets(file)
                for sheet, item in refs.items():
                    for cell, data in item['items'].items():
                        out.append({
                            'Sheet': sheet, 'Metric': data['names'],
                            'Cell': cell, 'Formula': data['formula'], 'References': data['references']
                        })
                        if verbose:
                            print(f"Sheet: {sheet}, Metrics: {data['names']}, Cell: {cell}, Formula: {data['formula']}, References: {data['references']}")

            txt_file.write(json.dumps(out, indent=2))
    else:
        with open(output_file + ".txt", "w+") as txt_file:
            out = []
            if sheet_name:
                refs = parse_single_sheet(file, sheet_name)
                for cell, data in refs.items():
                    out.append(f"Sheet: {sheet}, Metrics: {data['names']}, Cell: {cell}, Formula: {data['formula']}, References: {data['references']}\n")
                    if verbose:
                        print(f"Cell: {cell}, Formula: {data['formula']}, Refrences: {data['references']}")
            else:
                refs = parse_all_sheets(file)
                for sheet, item in refs.items():
                    for cell, data in item['items'].items():
                        out.append(f"Sheet: {sheet}, Metrics: {data['names']}, Cell: {cell}, Formula: {data['formula']}, References: {data['references']}\n")
                        if verbose:
                            print(f"Sheet: {sheet}, Metrics: {data['names']}, Cell: {cell}, Formula: {data['formula']}, References: {data['references']}")

            txt_file.writelines(out)

class SearchType(str, Enum):
    cell = "cell"
    sheet = "sheet"
    file = "file"
    file_sheet = "file-sheet"

@app.command()
def search(
    file: Annotated[str, typer.Argument(help="Path to the file.")],
    string: Annotated[str, typer.Argument(help="String to search for")],
    stype: Annotated[SearchType, typer.Option(case_sensitive=False)] = SearchType.cell,
):
    with open(file, "r") as txt:
        contents = json.loads(str(txt.read()))
        if (stype.value == SearchType.cell):
            print(json.dumps(search_cell(contents, string), indent=2))
        elif (stype.value == SearchType.file):
            print(json.dumps(search_ref_file(contents, string), indent=2))
        elif (stype.value == SearchType.sheet):
            print(json.dumps(search_ref_sheet(contents, string), indent=2))
        else:
            sstring = string.split(", ")
            print(json.dumps(search_ref_file_sheet(contents, sstring[0], sstring[1]), indent=2))


if __name__ == "__main__":
    app()
