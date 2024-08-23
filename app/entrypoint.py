from app.parse_tabular import main as tabular_main
from app.newform_main import entry as parse_main
# from app.main import main as parse_main
from app.make_summary import main as summary_main

if __name__ == "__main__":
    for fname in [
        "2024_Animal_DNA_Viruses_and_Retroviruses",
        "2024_Animal_dsRNA_and_ssRNA-_viruses",
        "2024_Archaeal_viruses",
        "2024_Bacterial_viruses",
        "2024_Fungal_and_Protist_Viruses",
        "2024_Plant_viruses"
    ]:
        # tabular_main(fname)
        parse_main(fname)
        summary_main(fname)