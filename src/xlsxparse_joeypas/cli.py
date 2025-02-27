import typer
import json
from xlsxparse_joeypas.parse import parse_all_sheets, parse_single_sheet
from typing_extensions import Annotated


def entry(
    file: Annotated[str, typer.Argument(help="Path to the .xlsx WorkBook")],
    sheet_name: Annotated[str, typer.Argument(help="Name of the sheet to parse (if none provided, parse all sheets)")] = "",
    to_json: Annotated[bool, typer.Option(help="Output file format as json.")] = False,
    verbose: Annotated[bool, typer.Option(help="Write output to console")] = False,
):
    if to_json:
        with open("output.json", "w+") as txt_file:
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
        with open("output.txt", "w+") as txt_file:
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



app = typer.Typer()
app.command()(entry)

if __name__ == "__main__":
    app()
