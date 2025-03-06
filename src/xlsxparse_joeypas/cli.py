import typer
import json
from xlsxparse_joeypas.parse import parse_all_sheets, parse_single_sheet
from xlsxparse_joeypas.search import search_ref_file, search_ref_sheet, search_ref_file_sheet, search_metric, search_sheet_metric
from typing_extensions import Annotated
from typing import Optional
from enum import Enum

__version__ = "0.1.2"

def version_callback(value: bool):
    if value:
        print(f"xlsxparse version: {__version__}")
        raise typer.Exit()

app = typer.Typer(no_args_is_help=True)

@app.command()
def create(
    file: Annotated[str, typer.Argument(help="Path to the .xlsx WorkBook")],
    sheet_name: Annotated[str, typer.Option("--sheet", help="Name of the sheet to parse (if none provided, parse all sheets)")] = None,
    verbose: Annotated[bool, typer.Option("--verbose", help="Write output to console")] = False,
    output_file: Annotated[str, typer.Option("--output", help="Path of output file")] = "output.json",
    version: Annotated[Optional[bool], typer.Option("--version", callback=version_callback)] = None,
):
    with open(output_file, "w+") as json_file:
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

        json_file.write(json.dumps(out, indent=2))


class SearchType(str, Enum):
    metric = "metric"
    sheet_metric = "sheet-metric"
    ref_file = "ref-file"
    ref_file_sheet = "ref-file-sheet"

@app.command()
def search(
    string: Annotated[str, typer.Argument(help="String to search for")],
    file: Annotated[str, typer.Argument(help="Path to the file.")] = "output.json",
    stype: Annotated[SearchType, typer.Option(case_sensitive=False)] = SearchType.metric,
):
    with open(file, "r") as txt:
        contents = json.loads(str(txt.read()))
        if (stype.value == SearchType.ref_file):
            print(json.dumps(search_ref_file(contents, string), indent=2))
        elif (stype.value == SearchType.metric):
            print(json.dumps(search_metric(contents, string), indent=2))
        elif (stype.value == SearchType.sheet_metric):
            sstring = string.split(", ")
            print(json.dumps(search_sheet_metric(contents, sstring[0], sstring[1]), indent=2))
        else:
            sstring = string.split(", ")
            print(json.dumps(search_ref_file_sheet(contents, sstring[0], sstring[1]), indent=2))


if __name__ == "__main__":
    app()
