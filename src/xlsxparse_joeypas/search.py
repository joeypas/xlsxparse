def search_ref_sheet(contents, sheet_name: str):
    return [obj for obj in contents if any(ref['sheet'] == sheet_name for ref in obj['References'])]

def search_ref_file(contents, file_name: str):
    return [obj for obj in contents if any("file" in ref and ref["file"] == file_name for ref in obj['References'])]

def search_ref_file_sheet(contents, file_name: str, sheet_name: str):
    return [obj for obj in contents if any("file" in ref and ref["file"] == file_name and ref["sheet"] == sheet_name for ref in obj['References'])]

def search_metric(contents, metric: str):
    return [obj for obj in contents if metric in obj["Metric"]]

def search_sheet_metric(contents, sheet_name: str, metric: str):
    return [obj for obj in contents if metric in obj["Metric"] and obj["Sheet"] == sheet_name]
