from report import ReportGenerator
import json
import os

def force_generate_reports():
    # Load settings to get default area
    settings = {}
    if os.path.exists("settings.json"):
        with open("settings.json", "r") as f:
            settings = json.load(f)
    
    def_area = settings.get("default_area", "")
    rg = ReportGenerator()
    
    print("Generating Weekly Report...")
    w = rg.generate_periodical_report('weekly', default_area=def_area)
    print(f"Result: {w}")
    
    print("Generating Monthly Report...")
    m = rg.generate_periodical_report('monthly', default_area=def_area)
    print(f"Result: {m}")
    
    print("Generating Yearly Report...")
    y = rg.generate_periodical_report('yearly', default_area=def_area)
    print(f"Result: {y}")

if __name__ == "__main__":
    force_generate_reports()
