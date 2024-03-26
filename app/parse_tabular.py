'''
Original script by @Donald Smith, ver 25/02/24
Adaptation by @Mayne941
'''

import re, csv, os
import pandas as pd
import pickle as p

def main():
    ExcelPath = "./data_tables/"
    CsvPath = f"./parsed_tables/"
    previous_type_change = ""
    Create_new_marker = "N"
    Renamed_marker = "N"
    Abolish_marker = "N"
    Move_marker = "N"
    MoveRename_marker = "N"
    Split_marker = "N"
    Merge_marker = "N"
    Demote_marker = "N"
    Promote_marker = "N"
    HTML_text = ""
    ignore_row = "N"
    code = "XXXX"
    expect_paired_lines = "N"
    split_taxon_name_1 = ""
    New_Excel_format = ""
    Excel_counter = -1

    New_taxon_table_starter_text = "><tr><th>Operation</th><th>Rank</th><th>New taxon name</th><th>Virus name</th><th>GenBank accession</th></tr>"
    New_taxon_table_text = ""
    Renamed_taxon_table_starter_text = "><tr><th>Operation</th><th>Rank</th><th>Previous taxon name</th><th>New taxon name</th></tr>"
    Renamed_taxon_table_text = ""
    Abolish_taxon_table_starter_text = "><tr><th>Operation</th><th>Rank</th><th>Abolished taxon name</th></tr>"
    Abolish_taxon_table_text = ""
    Move_taxon_table_starter_text = "><tr><th>Operation</th><th>Rank</th><th>Taxon name</th><th>Old parent taxon</th><th>New parent taxon</th></tr>"
    Move_taxon_table_text = ""
    MoveRename_taxon_table_starter_text = "><tr><th>Operation</th><th>Rank</th><th>Old taxon name</th><th>Old parent taxon</th><th>New taxon name</th><th>New parent taxon</th></tr>"
    MoveRename_taxon_table_text = ""
    Split_taxon_table_starter_text = "><tr><th>Operation</th><th>Rank</th><th>Old taxon</th><th>New taxon 1</th><th>Virus name</th><th>GenBank accession</th><th>New taxon 2</th><th>Virus name</th><th>GenBank accession</th></tr>"
    Split_taxon_table_text = ""
    Merge_taxon_table_starter_text = "><tr><th>Operation</th><th>Rank</th><th>Old taxon 1</th><th>Old taxon 2</th><th>Merged taxon</th></tr>"
    Merge_taxon_table_text = ""
    Demote_taxon_table_starter_text = "><tr><th>Operation</th><th>Old rank</th><th>Old taxon name</th><th>New rank</th><th>New taxon name</th></tr>"
    Demote_taxon_table_text = ""
    Promote_taxon_table_starter_text = "><tr><th>Operation</th><th>Old rank</th><th>Old taxon name</th><th>New rank</th><th>New taxon name</th></tr>"
    Promote_taxon_table_text = ""


    ranklist = ["species", "subgenus", "genus", "subfamily", "family", "suborder", "order", "subclass", "class",
                "subphylum", "phylum", "subkingdom", "kingdom", "subrealm", "realm"]

    # To do list of taxa for keywords
    # What to do if different operations within same table ...
    # Diacritics
    # Extracting code from top of sheet - doesn't work unless delete formula
    # All TP tables -  Move; rename, Split, Merge

    def zero_files(HTML_text, Create_new_marker, Renamed_marker, Abolish_marker, Move_marker, MoveRename_marker, Split_marker, Merge_marker, Demote_marker, Promote_marker, New_taxon_table_text, Renamed_taxon_table_text, Abolish_taxon_table_text, Move_taxon_table_text, MoveRename_taxon_table_text, Split_taxon_table_text, Merge_taxon_table_text, Demote_taxon_table_text, Promote_taxon_table_text ):
        if Create_new_marker == "Y":
            HTML_text = HTML_text + New_taxon_table_text
        elif Renamed_marker == "Y":
            HTML_text = HTML_text + Renamed_taxon_table_text
        elif Abolish_marker == "Y":
            HTML_text = HTML_text + Abolish_taxon_table_text
        elif Move_marker == "Y":
            HTML_text = HTML_text + Move_taxon_table_text
        elif MoveRename_marker == "Y":
            HTML_text = HTML_text + MoveRename_taxon_table_text
        elif Split_marker == "Y":
            HTML_text = HTML_text + Split_taxon_table_text
        elif Merge_marker == "Y":
            HTML_text = HTML_text + Merge_taxon_table_text
        elif Demote_marker == "Y":
            HTML_text = HTML_text + Demote_taxon_table_text
        elif Promote_marker == "Y":
            HTML_text = HTML_text + Promote_taxon_table_text
        #print("Function", HTML_text)
        return(HTML_text)

    fileList = [i for i in os.listdir(ExcelPath) if not "_Suppl" in i] # Supplements get downloaded with taxo forms
    if os.path.isfile(f"{CsvPath}output.csv"): os.remove(f"{CsvPath}output.csv")

    # Identify which Excel spreadsheet format is being used - assume all of same type
    New_Excel_format = input("2024 format (Y/N)? (So know which sheet to input)").lower()
    # New_Excel_format = "n"  #TODO RM HARD CODED FOR TEST

    for file in fileList:
        if New_Excel_format == "n":
            sheet = 0
        else:
            sheet = 1
        xls = pd.ExcelFile(f"{ExcelPath}{file}")
        df = pd.read_excel(xls, sheet_name=sheet)
        df.to_csv(f"{CsvPath}output.csv", mode="a", encoding="utf-8")       

    with open (f"{CsvPath}output.csv", encoding="utf-8") as csvfile:
        readcsv = csv.reader(csvfile, delimiter=',')
        for row in readcsv:
            try:
                if row[39] == "Please select":
                    continue
            except: breakpoint()

        
    # Identify start of TP sheet by presence of "Code:" and extract TP code from filename
            if re.search("Code", row[2]) or re.search("INSTRUCTIONS:", row[1]) or re.search("Code assigned", row[1]):
                Excel_counter = Excel_counter + 1
                code = fileList[Excel_counter]
                print(f"Processing file {Excel_counter}: {code}")
                #code = row[5]
                code=code[:9]
                ignore_row = "N"
                HTML_text = (zero_files(HTML_text, Create_new_marker, Renamed_marker, Abolish_marker, Move_marker,
                        MoveRename_marker, Split_marker, Merge_marker, Demote_marker, Promote_marker,
                        New_taxon_table_text, Renamed_taxon_table_text, Abolish_taxon_table_text,
                        Move_taxon_table_text, MoveRename_taxon_table_text, Split_taxon_table_text,
                        Merge_taxon_table_text, Demote_taxon_table_text, Promote_taxon_table_text))
                HTML_text = HTML_text + "</table><h2>Taxonomy proposal Code: " + code + "</h2>"
                Create_new_marker = "N"
                Renamed_marker = "N"
                Abolish_marker = "N"
                Move_marker = "N"
                MoveRename_marker = "N"
                Split_marker = "N"
                Merge_marker = "N"
                Demote_marker = "N"
                Promote_marker = "N"

                New_taxon_table_text = ""
                Renamed_taxon_table_text = ""
                Abolish_taxon_table_text = ""
                Move_taxon_table_text = ""
                MoveRename_taxon_table_text = ""
                Split_taxon_table_text = ""
                Merge_taxon_table_text = ""
                Demote_taxon_table_text = ""
                Promote_taxon_table_text = ""

                New_taxon_table_text = New_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + New_taxon_table_starter_text
                Renamed_taxon_table_text = Renamed_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Renamed_taxon_table_starter_text
                Abolish_taxon_table_text = Abolish_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Abolish_taxon_table_starter_text
                Move_taxon_table_text = Move_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Move_taxon_table_starter_text
                MoveRename_taxon_table_text = MoveRename_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + MoveRename_taxon_table_starter_text
                Split_taxon_table_text = Split_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Split_taxon_table_starter_text
                Merge_taxon_table_text = Merge_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Merge_taxon_table_starter_text
                Demote_taxon_table_text = Demote_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Demote_taxon_table_starter_text
                Promote_taxon_table_text = Promote_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Promote_taxon_table_starter_text

                continue

    # Identify heading rows - ignore these rows
            elif re.search("version", row[1]) or re.search("CURRENT TAXONOMY", row[1]) or re.search("CURRENT TAXONOMY", row[2]):
                continue

    #Identify column for species  in current and proposed taxonomy
            elif re.search("Realm", row[1]) or re.search("Realm", row[2]):
                current_species_column = 0
                proposed_species_column = 0
                for count, item in enumerate(row):
                    if re.search("Species", item):
                        if current_species_column == 0:
                            current_species_column = count
                        else:
                            proposed_species_column =  count
                continue

    # Extract information from taxonomic entry rows

    # Check if new type of change (to start new table)
            elif row[proposed_species_column + 8] != "":
                type_change = row[proposed_species_column + 8]
                if type_change != previous_type_change:
                    HTML_text = (zero_files(HTML_text, Create_new_marker, Renamed_marker, Abolish_marker, Move_marker, MoveRename_marker, Split_marker, Merge_marker, Demote_marker, Promote_marker, New_taxon_table_text, Renamed_taxon_table_text, Abolish_taxon_table_text, Move_taxon_table_text, MoveRename_taxon_table_text, Split_taxon_table_text, Merge_taxon_table_text, Demote_taxon_table_text, Promote_taxon_table_text))
                    Create_new_marker = "N"
                    Renamed_marker = "N"
                    Abolish_marker = "N"
                    Move_marker = "N"
                    MoveRename_marker = "N"
                    Split_marker = "N"
                    Merge_marker = "N"
                    Demote_marker = "N"
                    Promote_marker = "N"

                    New_taxon_table_text = ""
                    Renamed_taxon_table_text = ""
                    Abolish_taxon_table_text = ""
                    Move_taxon_table_text = ""
                    MoveRename_taxon_table_text = ""
                    Split_taxon_table_text = ""
                    Merge_taxon_table_text = ""
                    Demote_taxon_table_text = ""
                    Promote_taxon_table_text = ""

                    New_taxon_table_text = New_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + New_taxon_table_starter_text
                    Renamed_taxon_table_text = Renamed_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Renamed_taxon_table_starter_text
                    Abolish_taxon_table_text = Abolish_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Abolish_taxon_table_starter_text
                    Move_taxon_table_text = Move_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Move_taxon_table_starter_text
                    MoveRename_taxon_table_text = MoveRename_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + MoveRename_taxon_table_starter_text
                    Split_taxon_table_text = Split_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Split_taxon_table_starter_text
                    Merge_taxon_table_text = Merge_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Merge_taxon_table_starter_text
                    Demote_taxon_table_text = Demote_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Demote_taxon_table_starter_text
                    Promote_taxon_table_text = Promote_taxon_table_text + "<table style=\"width:100%\" border=\"1\" id=\"" + code + "\"><tbody" + Promote_taxon_table_starter_text

                    previous_type_change = type_change

    # Create new taxon option
                if re.search ("Create", row[proposed_species_column + 8]):
                    Create_new_marker = "Y"
                    changed_rank = row[proposed_species_column + 9]
                    for count, rank in enumerate (ranklist):
                        # if "Efunavirus" in row: breakpoint()
                        if changed_rank == rank:
                            new_taxon_name = row[proposed_species_column - count]
                            if changed_rank == "species":
                                exemplar = row[proposed_species_column - count + 2]
                                genbank = row[proposed_species_column - count + 1]
                            else:
                                exemplar = ""
                                genbank = ""
                            New_taxon_table_text = New_taxon_table_text + "<tr><td>New taxon</td><td>" + changed_rank + "</td><td><em>" + new_taxon_name + "</em></td><td>" + exemplar + "</td><td>" + genbank + "</td></tr>"
                            break

    # Rename option
                elif row[proposed_species_column + 8] == "Rename":
                    Renamed_marker = "Y"
                    changed_rank = row[proposed_species_column + 9]
                    for count, rank in enumerate(ranklist):
                        if changed_rank == rank:
                            new_taxon_name = row[proposed_species_column - count]
                            old_taxon_name = row[current_species_column - count]
                            Renamed_taxon_table_text = Renamed_taxon_table_text + "<tr><td>Rename taxon</td><td>" + changed_rank + "</td><td><em>" + old_taxon_name + "</em></td><td><em>" + new_taxon_name + "</em></td></tr>"
                            break

    # Abolish option
                elif row[proposed_species_column + 8] == "Abolish":
                    Abolish_marker = "Y"
                    changed_rank = row[proposed_species_column + 9]
                    for count, rank in enumerate(ranklist):
                        if changed_rank == rank:
                            old_taxon_name = row[current_species_column - count]
                            Abolish_taxon_table_text = Abolish_taxon_table_text + "<tr><td>Abolish taxon</td><td>" + changed_rank + "</td><td><em>" + old_taxon_name + "</em></td></tr>"
                            break

    # Move rename
                elif re.search("rename", row[proposed_species_column + 8]):
                    MoveRename_marker = "Y"
                    changed_rank = row[proposed_species_column + 9]
                    for count, rank in enumerate(ranklist):
                        if changed_rank == rank:
                            old_taxon_name = row[current_species_column - count]
                            new_taxon_name = row[proposed_species_column - count]
                            for x in range(15):
                                if row[current_species_column - 14 + x] != row[proposed_species_column - 14 + x]:
                                    if row[current_species_column - 14 + x] != "":
                                        old_parent_taxon_name = row[current_species_column - 14 + x]
                                        new_parent_taxon_name = row[proposed_species_column - 14 + x]
                    # This picks up the case where new parent taxon not present in previous taxonomy
                                    else:
                                        old_parent_taxon_name = row[current_species_column - 14 + x -1]
                                        new_parent_taxon_name = row[proposed_species_column - 14 + x]
                                    changed_parent_taxon_rank = ranklist[x] # RM < TODO 264-5 INDENTED 1
                                    MoveRename_taxon_table_text = MoveRename_taxon_table_text + "<tr><td>Move; rename taxon</td><td>" + changed_rank + "</td><td><em>" + old_taxon_name + "</em></td><td><em>" + old_parent_taxon_name + "</em></td><td><em>" + new_taxon_name + "</em></td><td><em>" + new_parent_taxon_name + "</em></td></tr>"
                                    break
                            break

    # Move option
                elif row[proposed_species_column + 8] == "Move":
                    Move_marker = "Y"
                    changed_rank = row[proposed_species_column + 9]
                    for count, rank in enumerate(ranklist):
                        if changed_rank == rank:
                            taxon_name = row[proposed_species_column - count]
                            for x in range (15):
                                if row[current_species_column - 14 + x] != row[proposed_species_column - 14 + x]:
                                    if row[current_species_column - 14 + x] != "":
                                        old_parent_taxon_name = row[current_species_column - 14 + x]
                                        new_parent_taxon_name = row[proposed_species_column - 14 + x]
                                    else:
                                        old_parent_taxon_name = row[current_species_column - 14 + x - 1]
                                        new_parent_taxon_name = row[proposed_species_column - 14 + x]
                                    Move_taxon_table_text = Move_taxon_table_text + "<tr><td>Move taxon</td><td>" + changed_rank + "</td><td><em>" + taxon_name + "</em></td><td><em>" + old_parent_taxon_name + "</em></td><td><em>" + new_parent_taxon_name + "</em></td></tr>"
                                    break
                            break

    # Promote option
                elif row[proposed_species_column + 8] == "Promote":
                    Promote_marker = "Y"
                    changed_rank = row[proposed_species_column + 9]
                    for count, rank in enumerate(ranklist):
                        if changed_rank == rank:
                            taxon_name = row[current_species_column - count]
                            for x in range(1, 15 - count):
                                if row[current_species_column - 14 + x] != row[proposed_species_column - 14 + x]:
                                    promoted_taxon_name = row[proposed_species_column - 14 + x]
                                    Promote_taxon_table_text = Promote_taxon_table_text + "<tr><td>Promote taxon</td><td>" + changed_rank + "</td><td><em>" + taxon_name + "</em></td><td>" + \
                                                            ranklist[14 - x] + "</td><td><em>" + promoted_taxon_name + "</em></td></tr>"
                                    break
                            break

    # Demote option
                elif row[proposed_species_column + 8] == "Demote":
                    Demote_marker = "Y"
                    changed_rank = row[proposed_species_column + 9]
                    for count, rank in enumerate(ranklist):
                        if changed_rank == rank:
                            taxon_name = row[current_species_column - count]
                            for x in range(1, 15 - count):
                                if row[current_species_column - x] != row[proposed_species_column - x]:
                                    demoted_taxon_name = row[proposed_species_column - x]
                                    Demote_taxon_table_text = Demote_taxon_table_text + "<tr><td>Demote taxon</td><td>" + changed_rank + "</td><td><em>" + taxon_name + "</em></td><td>" + \
                                                            ranklist[x] + "</td><td><em>" + demoted_taxon_name + "</em></td></tr>"
                                    break
                            break

    # Split option
                elif row[proposed_species_column + 8] == "Split":
                    Split_marker = "Y"
                    changed_rank = row[proposed_species_column + 9]
                    if expect_paired_lines == "N":
                        split_taxon_name_1 = ""
                        expect_paired_lines = "Y"
                        for count, rank in enumerate(ranklist):
                            if changed_rank == rank:
                                old_taxon_name = row[current_species_column - count]
                                split_taxon_name_1 = row[proposed_species_column - count]
                                exemplar_1 = row[proposed_species_column + 2]
                                genbank_1 = row[proposed_species_column + 1]
                                break
                    else:
                        for count, rank in enumerate(ranklist):
                            if changed_rank == rank:
                                old_taxon_name = row[current_species_column - count]
                                split_taxon_name_2 = row[proposed_species_column - count]
                                exemplar_2 = row[proposed_species_column + 2]
                                genbank_2 = row[proposed_species_column + 1]
                                Split_taxon_table_text = Split_taxon_table_text + "<tr><td>Split taxon</td><td>" + changed_rank + "</td><td><em>" + old_taxon_name + "</em></td><td><em>" + split_taxon_name_1 + "</em></td><td>" + exemplar_1 + "</td><td>" + genbank_1 + "</td><td><em>" + split_taxon_name_2 + "</em></td><td>" + exemplar_2 + "</td><td>" + genbank_2 + "</td></tr>"
                                split_taxon_name_1 = ""
                                expect_paired_lines = "N"
                                exemplar_1 = ""
                                genbank_1 = ""
                                break

    # Merge option
                elif row[proposed_species_column + 8] == "Merge":
                    Merge_marker = "Y"
                    changed_rank = row[proposed_species_column + 9]
                    if expect_paired_lines == "N":
                        merged_taxon_name = ""
                        expect_paired_lines = "Y"
                        for count, rank in enumerate(ranklist):
                            if changed_rank == rank:
                                old_taxon_name_1 = row[current_species_column - count]
                                merged_taxon_name = row[proposed_species_column - count]
                                break
                    else:
                        for count, rank in enumerate(ranklist):
                            if changed_rank == rank:
                                old_taxon_name_2 = row[current_species_column - count]
                                Merge_taxon_table_text = Merge_taxon_table_text + "<tr><td>Merge taxa</td><td>" + changed_rank + "</td><td><em>" + old_taxon_name_1 + "</em></td><td><em>" + old_taxon_name_2 + "</em></td><td><em>" + merged_taxon_name + "</em></td></tr>"
                                expect_paired_lines = "N"
                                old_taxon_name_1 = ""
                                break
            # if "Efunavirus" in row: breakpoint() ##############

        #Adds on last table
        HTML_text = (zero_files(HTML_text, Create_new_marker, Renamed_marker, Abolish_marker, Move_marker, MoveRename_marker,
                Split_marker, Merge_marker, Demote_marker, Promote_marker, New_taxon_table_text,
                Renamed_taxon_table_text, Abolish_taxon_table_text, Move_taxon_table_text, MoveRename_taxon_table_text,
                Split_taxon_table_text, Merge_taxon_table_text, Demote_taxon_table_text, Promote_taxon_table_text))

        HTML_text = HTML_text + "</table>"
        HTML_text = HTML_text[8:]

        with open("combined_tables_html.txt", "w", encoding="'utf-8'") as output:
            output.write(HTML_text)

        master_df = pd.DataFrame()
        tables_split = HTML_text.split("<h2>")
        del tables_split[0] # kill empty
        for table_block in tables_split:
            code = table_block.split("</h2>")[0].split(" ")[-1]
            tables = table_block.split("<tbody>")
            del tables[0]
            for table in tables:
                table = f"<tbody>{table}"
                try:
                    tmp_df = pd.read_html(table)
                except ValueError: 
                    tmp_df = pd.read_html(f"<table>{table}")
                tmp_df[0]["tp"] = code
                master_df = pd.concat([master_df, tmp_df[0]])

        # breakpoint()
        # df = pd.read_html(HTML_text)
        # for idx, subtable in enumerate(df):
        #     try:
        #         subtable["tp"] = fileList[idx].split(".xlsx")[0]
        #     except: breakpoint()
        #     master_df = pd.concat([master_df, subtable])

        pickle_fname = "animal_+ssrna.p"
        with open(pickle_fname, "wb") as f:
            p.dump(master_df, f)

if __name__ == "__main__":
    main()