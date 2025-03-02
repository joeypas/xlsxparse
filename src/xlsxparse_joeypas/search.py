def search_cell(contents, cell: str):
    return [x for x in j if x["Cell"] == cell]

def search_ref_sheet(contents, sheet_name: str):
    return [obj for obj in contents if any(ref['sheet'] == sheet_name for ref in obj['References'])]

def search_ref_file(contents, file_name: str):
    return [obj for obj in contents if any("file" in ref and ref["file"] == file_name for ref in obj['References'])]

def search_ref_file_sheet(contents, file_name: str, sheet_name: str):
    return [obj for obj in contents if any("file" in ref and ref["file"] == file_name and ref["sheet"] == sheet_name for ref in obj['References'])]

