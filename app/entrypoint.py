from app.parse_tabular import main as tabular_main
from app.newform_main import entry as parse_main
# from app.main import main as parse_main
from app.make_summary import main as summary_main

if __name__ == "__main__":
    fname = "2024_plant_viruses"
    tabular_main(fname)
    parse_main(fname)
    summary_main(fname)