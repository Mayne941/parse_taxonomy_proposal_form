from app.parse_tabular import main as tabular_main
from app.newform_main import entry as parse_main
from app.make_summary import main as summary_main

if __name__ == "__main__":
    '''DATASET FOLDERS - SCRIPT EXPECTS THESE IN REPOSITORY BASE DIR'''
    for fname in [ 
        "2026_Animal_DNA_Viruses_and_Retroviruses",
        "2026_Animal_dsRNA_and_ssRNA-_viruses", 
        "2026_Archaeal_viruses",
        "2026_Bacterial_viruses", 
        "2026_Fungal_and_Protist_Viruses",
        '2026_Animal_ssRNA+_viruses',
        "2026_Plant_viruses"
    ]:
        tabular_main(fname)
        parse_main(fname)
        summary_main(fname)