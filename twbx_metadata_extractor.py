import zipfile, os, shutil, re
from bs4 import BeautifulSoup
import pandas as pd

def extract_twb_from_twbx(twbx_path):
    with zipfile.ZipFile(twbx_path, 'r') as z:
        for file in z.namelist():
            if file.endswith('.twb'):
                z.extract(file, path="temp_extracted")
                return os.path.join("temp_extracted", file)
    raise FileNotFoundError("No .twb file found in .twbx archive.")

def simplify_federated_name(raw_field):
    if not raw_field:
        return None

    # Remove outer brackets if present
    raw_field = raw_field.strip('[]')

    # Match inner part: [federated.xxx].[prefix:FieldName:maybeAlias]
    # Extract prefix and field name
    match = re.search(r'\.\[(?:mn|attr|agg|dim|meas|none):([^:\]]+)', raw_field)
    if match:
        return match.group(1)

    # If not matched, fallback: try to extract anything after last colon
    fallback = re.findall(r':([^:\]]+)\]?', raw_field)
    if fallback:
        return fallback[-1]

    # Final fallback: remove IDs and return raw
    return re.sub(r'federated\.[^.]+\.', '', raw_field)

def resolve_friendly_name(raw_field, friendly_names):
    simplified = simplify_federated_name(raw_field)
    return friendly_names.get(raw_field) or friendly_names.get(simplified) or simplified or raw_field

def extract_field_references(formula):
    """Extract field references from a formula using regex."""
    return re.findall(r'\[.*?\]', formula) if formula else []

def parse_twb(twb_path):
    with open(twb_path, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f, 'xml')

    metadata = {
        "workbook_name": soup.workbook.get("name"),
        "worksheets": [],
        "dashboards": [],
        "data_sources": [],
        "fields": [],
        "calculated_fields": [],
        "parameters": [],
        "lineage": []
    }

    # Worksheets
    # First collect reader friendly names for shelf data in worksheets
    friendly_names = {}
    for col in soup.find_all('column'): 
        raw = col.get('name') 
        label = col.get('caption') or col.get('alias') or raw 
        friendly_names[raw] = label
    
    friendly_names = {}
    for col in soup.find_all('column'):
        internal_name = col.get('name')
        caption = col.get('caption') or col.get('alias') or simplify_federated_name(internal_name)
        friendly_names[internal_name] = caption
        simplified = simplify_federated_name(internal_name)
        if simplified != internal_name:
            friendly_names[simplified] = caption

    # Collect worksheet data
    for ws in soup.find_all("worksheet"):
        ws_name = ws.get("name")
        row_fields, col_fields, filter_fields = [], [], []
        # Collect shelf data
        for shelf in ws.find_all('shelf'):
            shelf_name = shelf.get('name', '').lower()

            # Match shelves flexibly
            if 'row' in shelf_name:
                target = row_fields
            elif 'column' in shelf_name:
                target = col_fields
            else:
                target = None

            # Parse <field> entries
            if target is not None:
                for field in shelf.find_all(['field', 'tableau-field']):
                    raw_name = field.get('name')
                    if raw_name:
                        clean = resolve_friendly_name(raw_name, friendly_names)
                        target.append(clean)

                # Parse <encoding><expression> entries
                for encoding in shelf.find_all('encoding'):
                    for expr in encoding.find_all('expression'):
                        if expr.get('class') == 'column':
                            raw_expr = expr.get('value')
                            clean = resolve_friendly_name(raw_expr, friendly_names)
                            target.append(clean)

        for flt in ws.find_all('filter'):
            field_name = flt.get('field') or flt.get('column')
            clean_name = simplify_federated_name(field_name)
            readable_name = friendly_names.get(clean_name, clean_name)
            expression = flt.get('expression') or str(flt.attrs)
            if readable_name:
                filter_fields.append(readable_name)

        metadata["worksheets"].append({
            'name': ws_name,
            'Rows': row_fields,
            'Columns': col_fields,
            'Filters': filter_fields,
            'Filter_logic': expression
            })

    # Dashboards
    for db in soup.find_all("dashboard"):
        metadata["dashboards"].append(db.get("name"))

    # Data Sources and Fields
    for ds in soup.find_all("datasource"):
        ds_info = {
            "name": ds.get("name"),
            "caption": ds.get("caption"),
            "connections": [conn.attrs for conn in ds.find_all('connection')],
            "custom_sql": [rel.get('text') for rel in ds.find_all('relation', {'type': 'text'})]
        }
        for col in ds.find_all("column"):
            field = {
                "name": col.get("name"),
                "caption": col.get("caption"),
                "datatype": col.get("datatype"),
                "role": col.get("role"),
                "type": col.get("type"),
                "formula": None,
                "is_calculated": False,
                "datasource": ds_info["name"]
            }
            calc = col.find("calculation")
            if calc:
                field["formula"] = calc.get("formula")
                field["is_calculated"] = True
                refs = extract_field_references(field["formula"])
                metadata["calculated_fields"].append({
                    "field_name": field["name"],
                    "caption": field["caption"],
                    "formula": field["formula"],
                    "datasource": ds_info["name"],
                    "references": refs
                })
                for ref in refs:
                    metadata["lineage"].append({
                        "calculated_field": field["name"],
                        "depends_on": ref,
                        "datasource": ds_info["name"]
                    })
            #ds_info["fields"].append(field)
            metadata["fields"].append(field)
        metadata["data_sources"].append(ds_info)

    # Parameters
    for param in soup.find_all("datasource"):
        if param.get('name') == 'Parameters':
            for col in param.find_all('column'):
                metadata["parameters"].append({
                    "name": col.get("name"),
                    "caption": col.get("caption"),
                    "data_type": col.get("datatype"),
                    'alias': col.get('alias'),
                    'value': col.get('value'),
                    "current_value": col.get("currentValue")
                })

    return metadata

def cleanup_temp():
    shutil.rmtree("temp_extracted", ignore_errors=True)

def export_to_excel(metadata, output_file="tableau_metadata.xlsx"):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        pd.DataFrame(metadata["data_sources"]).to_excel(writer, sheet_name="Datasources", index=False)
        pd.DataFrame(metadata["worksheets"]).to_excel(writer, sheet_name="Worksheets", index=False)
        pd.DataFrame(metadata["dashboards"], columns=["Dashboard Name"]).to_excel(writer, sheet_name="Dashboards", index=False)
        pd.DataFrame(metadata["parameters"]).to_excel(writer, sheet_name="Parameters", index=False)
        pd.DataFrame(metadata["fields"]).to_excel(writer, sheet_name="Fields", index=False)
        pd.DataFrame(metadata["calculated_fields"]).to_excel(writer, sheet_name="Calculated Fields", index=False)
        pd.DataFrame(metadata["lineage"]).to_excel(writer, sheet_name="Lineage", index=False)

def extract_tableau_metadata(file_path):
    try:
        twb_path = extract_twb_from_twbx(file_path) if file_path.endswith(".twbx") else file_path
        metadata = parse_twb(twb_path)
        cleanup_temp()
        export_to_excel(metadata)
        print("Metadata exported to tableau_metadata.xlsx")
    except Exception as e:
        print(f"Error: {e}")

# === Run the script ===
if __name__ == "__main__":
    tableau_file_path = "Your twb or twbx"
    extract_tableau_metadata(tableau_file_path)  # or .twb