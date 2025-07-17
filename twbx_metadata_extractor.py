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
    for ws in soup.find_all("worksheet"):
        ws_name = ws.get("name")
        metadata["worksheets"].append(ws_name)

    # Dashboards
    for db in soup.find_all("dashboard"):
        metadata["dashboards"].append(db.get("name"))

    # Data Sources and Fields
    for ds in soup.find_all("datasource"):
        ds_info = {
            "name": ds.get("name"),
            "caption": ds.get("caption"),
            "fields": []
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
            ds_info["fields"].append(field)
            metadata["fields"].append(field)
        metadata["data_sources"].append(ds_info)

    # Parameters
    for param in soup.find_all("parameter"):
        metadata["parameters"].append({
            "name": param.get("name"),
            "data_type": param.get("dataType"),
            "current_value": param.get("currentValue")
        })

    return metadata

def cleanup_temp():
    shutil.rmtree("temp_extracted", ignore_errors=True)

def export_to_excel(metadata, output_file="tableau_metadata.xlsx"):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        pd.DataFrame(metadata["worksheets"], columns=["Worksheet Name"]).to_excel(writer, sheet_name="Worksheets", index=False)
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
        print("✅ Metadata exported to tableau_metadata.xlsx")
    except Exception as e:
        print(f"❌ Error: {e}")

# === Run the script ===
if __name__ == "__main__":
    extract_tableau_metadata("your_workbook.twbx")  # or .twb