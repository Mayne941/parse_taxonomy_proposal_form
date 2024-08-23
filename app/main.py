from bs4 import BeautifulSoup
import zipfile
import json 
import os

class Parse:
    def __init__(self, fpath) -> None:
        self.fpath = fpath
        self.raw_data = []
        self.attribs = {}

    def scrape(self):
        '''Read word file in, convert to XML then stucture as dict'''
        document = zipfile.ZipFile(self.fpath)
        xml_content = document.read('word/document.xml')
        document.close()
        soup = BeautifulSoup(xml_content, 'xml')

class Strip:
    def __init__(self, fname, in_dir, out_dir, do_optional) -> None:
        self.fname = fname
        self.in_dir = in_dir
        self.out_dir = out_dir
        self.do_optional = do_optional
        self.attribs, self.parser_errors = {}, {}
        self.which_section = 2
        self.parser_err_codes = ["PARSER_ERROR", "MISSING_FIELD"]
        self.essential_fields = [
            "Title", "Id_code", "Authors", "Corr_author", "Subcommittees", "Study_groups", "Submission_date", "Tp_abstract", "Tp_type", "Revision_date"
        ]
        self.section_fns = {
            "Title": self.populate_title,
            "Code assigned:": self.populate_id,
            "Author(s), affiliation and email address(es) –": self.populate_authors,
            "Corresponding author(s)": self.get_main_author,
            "Sub-committee": self.get_subcommittee,
            "List the ICTV Study Group(s) that have seen or who have involved in creating this proposal.": self.get_study_groups,
            "Submission date": self.get_subm_date,
            "Optional – complete if formally voted on by an ICTV Study Group": self.populate_group_vote,
            "Executive Committee Meeting Decision code": self.get_meeting_decision,
            "Comments from the Executive Committee": self.get_comments,
            "Response of proposer": self.get_response,
            "Revision date": self.get_rev_date,
            "Abstract": self.get_abstract,
            "Text of General Proposal": self.get_general_proposal,
            "References": self.get_references,
            "Text of Taxonomy proposal": self.get_taxonomy_proposal,
            "Name of accompanying Excel module": self.get_excel_name,
            "Taxonomic changes proposed": self.get_proposed_changes,
            "Is any taxon name used here derived from that of a living person (Y/N)": self.get_vanity_names
        }

        for doc in raw_contents:
            '''Fix excel path'''
            for item in doc:
                item = item.replace(".xlxs", ".xlsx")
        breakpoint()

        self.raw_data = raw_contents

    def get_metadata(self):
        '''Get id code and title'''
        datum = self.raw_data[0]
        self.attribs["code"] = datum[[i for i, x in enumerate(datum) if "Code assigned:" in x][0] + 1]
        if self.attribs["code"] == "": breakpoint()
        self.attribs["title"] = datum[[i for i, x in enumerate(datum) if "Short title:" in x][0]].replace("Short title: ", "")
        self.attribs["study_grp"] = self.raw_data[4][0]

    def get_authors(self):
        '''Scrape author details'''
        authors = [i for i in self.raw_data[1] if not i == ""]
        addresses = self.raw_data[2]
        if authors == "": breakpoint()
        self.attribs["authors"] = {}
        self.attribs["authors"]["names"] = [i.strip().replace(".", "") for i in authors[0].replace(";", ",").split(",") if not i.strip().replace(" ", "") == ""]
        self.attribs["authors"]["emails"] = [i.strip() for i in authors[1].replace(";", ",").split(",") if not i.strip().replace(" ", "") == ""]
        self.attribs["authors"]["addresses"] = [f"{i}" for i in addresses[0].replace("]",")").replace("[","").replace(")",")@~").split("@~") if not i.strip().replace(" ", "") == ""]
        self.attribs["authors"]["corr_author"] = self.raw_data[3][0]

        # if self.attribs["code"] == "<i>2023.027B</i>": breakpoint()

    def get_content(self, submission_idx, excel_idx):
        '''Attempt to scrape content from floating boxes.
        Indices of boxes vary depending on filled, unfilled and deleted boxes.
        Attempt to center it by finding submission date and excel box with some string matching.
        '''
        try:
            self.attribs["submission_date"] = self.raw_data[submission_idx][1]
        except: 
            submission_idx -= 1
            try:
                self.attribs["submission_date"] = self.raw_data[submission_idx][1]
            except:
                self.attribs["submission_date"] = self.raw_data[submission_idx]
        try:
            if self.attribs["submission_date"].lower() == "person from whom the name is derived":
                submission_idx += 1
                self.attribs["submission_date"] = self.raw_data[submission_idx][1]
            elif self.attribs["submission_date"].lower() == "n" or self.attribs["submission_date"].lower() == "y" or ".xlsx" in self.attribs["submission_date"] or self.attribs["submission_date"].lower() == "no":
                self.attribs["submission_date"] = self.raw_data[submission_idx+2][1]
                submission_idx += 2
            elif self.attribs["submission_date"].lower() == "number of members":
                self.attribs["submission_date"] = self.raw_data[submission_idx+3][1]
                submission_idx += 3
        except:
            self.attribs["submission_date"] = "<<COULDN'T PARSE SUBMISSION DATE>>"

        # if self.attribs["submission_date"] == "Person from whom the name is derived":breakpoint()

        try:
            self.attribs["revision_date"] = self.raw_data[submission_idx][3]
        except:
            self.attribs["revision_date"] = "<<COULDN'T PARSE SUBMISSION DATE>>"

        if ".xl" in self.raw_data[excel_idx][0]:
            excel_box = 8
        elif ".xl" in self.raw_data[excel_idx+1][0]:
            excel_box = 9
        elif ".xl" in self.raw_data[excel_idx+2][0]:
            excel_box = 10
        elif ".xl" in self.raw_data[excel_idx+3][0]:
            excel_box = 11
        elif ".xl" in self.raw_data[excel_idx+4][0]:
            excel_box = 12
        else: breakpoint()

        self.attribs["excel_fname"] = self.raw_data[excel_box][0]
        try:
            self.attribs["abstract"] = self.raw_data[excel_box+1][0]
        except: breakpoint()
        self.attribs["proposal_text"] = " ".join(self.raw_data[excel_box+2])

    def main(self):
        self.scrape()
        self.get_metadata()
        self.get_authors()

    def get_study_groups(self, _row, _cell_idx, row_idx, table) -> None:
        '''Get study group/s'''
        try:
            groups = [i.text for i in table.rows[row_idx + 1].cells if not i.text.strip().replace(" ","") == ""]
            assert groups != []
            self.attribs["Study_groups"] = groups
        except:
            self.attribs["Study_groups"] = self.parser_errors["Study_groups"] = self.parser_err_codes[1]

    def get_subm_date(self, row, cell_idx, *_) -> None:
        '''Get submission date'''
        try:
            subm_date = [i.text for i in row.cells[cell_idx+1].paragraphs[0].runs if not i.text.strip().replace(" ","") == ""]
            assert subm_date != []
            self.attribs["Submission_date"] = subm_date 
        except:
            self.attribs["Submission_date"] = self.parser_errors["Submission_date"] = self.parser_err_codes[1]

    def populate_group_vote(self, _row, _cell_idx, row_idx, table) -> None:
        '''Optional: get group vote numbers'''
        if self.do_optional:
            group_vote_responses = []
            counter = 3 # first blank row
            while True:
                try:
                    row_text = [i.text for i in table.rows[row_idx + counter].cells]
                except IndexError:
                    '''End of table'''
                    break
                counter += 1
                if not row_text == ["", "", "", ""]:
                    for cell in row_text:
                        if cell == "":
                            cell = 0
                    group_vote_responses.append({"group": row_text[0], "support": row_text[1], "against": row_text[2], "no vote": row_text[3]})
            self.attribs["Study_group_votes"] = group_vote_responses

    def get_meeting_decision(self, _row, _cell_idx, row_idx, table) -> None:
        '''Optional: get decision of subcommittee meeting'''
        if self.do_optional:
            decision = []
            counter = 1
            while True:
                try:
                    row_text = [i.text for i in table.rows[row_idx + counter].cells]
                except IndexError:
                    '''End of table'''
                    break
                counter += 1
                if row_text[1] != "":
                    decision.append(row_text[0])
            self.attribs["Ex_committee_decision"] = decision

    def get_comments(self, _row, _cell_idx, row_idx, table) -> None:
        '''Optional: get author comments'''
        if self.do_optional:
            comments = []
            counter = 1
            while True:
                try:
                    _ = [i.text for i in table.rows[row_idx + counter].cells]
                except IndexError:
                    '''End of table'''
                    break
                counter += 1

                for para in table.rows[row_idx + 1].cells[0].paragraphs:
                    for i in para.runs:
                        text = i.text 
                        its = i.font.italic
                        if its:
                            text = f"<i>{text}<\i>"
                        comments.append(text)
            self.attribs["Ex_committee_comments"] = comments

    def get_response(self, _row, _cell_idx, row_idx, table) -> None:
        '''Optional: get subcommittee comments'''
        if self.do_optional:
            response = []
            counter = 1
            while True:
                try:
                    _ = [i.text for i in table.rows[row_idx + counter].cells]
                except IndexError:
                    '''End of table'''
                    break
                counter += 1
                for para in table.rows[row_idx + 1].cells[0].paragraphs:
                    for i in para.runs:
                        text = i.text 
                        its = i.font.italic
                        if its:
                            text = f"<i>{text}<\i>"
                        response.append(text)
            self.attribs["Proposer_response"] = response

    def get_rev_date(self, row, cell_idx, *_) -> None:
        '''Get revision date'''
        try:
            rev_date = [i.text for i in row.cells[cell_idx+1].paragraphs[0].runs if not i.text.strip().replace(" ","") == ""]
            assert rev_date != []
            self.attribs["Revision_date"] = rev_date
        except:
            self.attribs["Revision_date"] = self.parser_errors["Revision_date"] = self.parser_err_codes[1]

    def get_abstract(self, _row, _cell_idx, row_idx, table) -> None:
        '''Parse abstract from either S2 or S3'''
        try:
            abstract = []
            if [i.text for i in table.rows[row_idx +1].cells] == ['Brief description of current situation:       \n\n\nProposed changes:     \n\n\nJustification:      \n\n']:
                # TODO I've guessed what a blank box looks like: needs to be tested + made more robust
                return 
            for para in table.rows[row_idx + 1].cells[0].paragraphs:
                for i in para.runs:
                    text = i.text 
                    its = i.font.italic
                    if its:
                        text = f"<i>{text}<\i>"
                    abstract.append(text)

            '''Make flag to indicate whether sec 2 or 3 was filled in'''
            if self.which_section == 2:
                self.attribs["Tp_type"] = ["Non-taxonomic proposal"]

            elif self.which_section == 3:
                assert not "Tp_type" in self.attribs.keys(), "Error: User has filled in both section 2 + 3"
                assert abstract != []
                self.attribs["Tp_type"] = ["Taxonomic proposal"]
            self.attribs["Tp_abstract"] = abstract      
        except:
              self.attribs["Tp_abstract"] = self.parser_errors["Tp_abstract"] = self.parser_err_codes[0]

    def collate_errors(self) -> str:
        '''Create pickle object containing details of errors, if present'''
        errors = {}
        for field in self.essential_fields:
            '''Mark absent essential fields'''
            if not field in self.attribs.keys():
                if not "missing_fields" in errors.keys():
                    errors["missing_fields"] = []
                errors["missing_fields"].append(field)
        errors = {**errors, **self.parser_errors}

        if errors:
            hash = r.getrandbits(128)
            errors["document_name"] = self.fname 
            if not "Id_code" in self.attribs.keys():
                self.attribs["Id_code"] = [f"Not_defined{dt.datetime.now().strftime('%Y%M%d%H%m%s')}"]
            errors["Id_code"] = self.attribs["Id_code"]
            pickle.dump(errors, open(f"{hash}.pickle", "wb"))
            return f"{hash}.pickle"
        else:
            return "na"
        
    def save_json(self) -> None:
        '''Dump results to machine-readable format'''
        with open(f"{self.out_dir}{self.attribs['Id_code'][0]}.json", "w") as outfile: 
            json.dump(self.attribs, outfile)                

    def make_summary(self) -> None:
        '''Dump results to docx file'''
        doc = Document()
        for title, content in self.attribs.items():
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1
            p.paragraph_format.space_after = 0
            run_header = p.add_run(f"{title}\n")
            run_header.bold = True
            for cont_block in content:
                if "<i>" in cont_block:
                    run = p.add_run(cont_block.replace("<i>", "").replace("<\i>",""))
                    run.italic = True
                else:
                    run = p.add_run(cont_block)
                    run.italic = False   
                if title == "Authors":
                    p.add_run("\n")
                    if "@" in cont_block:
                        p.add_run("\n")
                    
            p.add_run("\n")
        doc.save(f"{self.out_dir}{self.attribs['Id_code'][0]}.docx")

    def get_general_proposal(self, _row, _cell_idx, row_idx, table):
        if do_optional:
            try: # TODO Lots of similarity with abstract box; try to combine
                proposal = []
                if [i.text.replace("Background:","").replace("Proposed changes:","").replace("Justification:","").replace("\n", "").strip().replace(" ","") for i in table.rows[row_idx +1].cells] == ['']:
                    # RM < TODO UPDATE REPLACES WITH NEW FORM FORMAT
                    return 
                for para in table.rows[row_idx + 1].cells[0].paragraphs:
                    for i in para.runs:
                        text = i.text 
                        its = i.font.italic
                        if its:
                            text = f"<i>{text}<\i>"
                        proposal.append(text)

                '''Make flag to indicate whether sec 2 or 3 was filled in'''
                assert proposal != []
                self.attribs["general_proposal"] = proposal      
            except:
                self.attribs["general_proposal"] = self.parser_errors["general_proposal"] = self.parser_err_codes[0]

    def get_references(self, _row, _cell_idx, row_idx, table):
        if do_optional:
            try: # TODO Lots of similarity with abstract box; try to combine
                refs = []
                if [i.text.replace("\n", "").strip().replace(" ","") for i in table.rows[row_idx +1].cells] == ['']:
                    return 
                for para in table.rows[row_idx + 1].cells[0].paragraphs:
                    for i in para.runs:
                        text = i.text 
                        its = i.font.italic
                        if its:
                            text = f"<i>{text}<\i>"
                        refs.append(text)

                '''Make flag to indicate whether sec 2 or 3 was filled in'''
                assert refs != []
                self.attribs["references"] = refs      
            except:
                self.attribs["references"] = self.parser_errors["references"] = self.parser_err_codes[0]

    def get_taxonomy_proposal(self, _row, _cell_idx, row_idx, table):
        if do_optional:
            try: # TODO Lots of similarity with abstract box; try to combine
                tp_text = []
                if [i.text.replace("Taxonomic level(s) affected:","").replace("Description of current taxonomy:","").replace("Proposed taxonomic change(s):","").replace("Justification:","").replace("\n", "").strip().replace(" ","") for i in table.rows[row_idx +1].cells] == ['']:
                    return # RM < TODO UPDATE REPLACES WITH NEW FORM FORMAT
                for para in table.rows[row_idx + 1].cells[0].paragraphs:
                    for i in para.runs:
                        text = i.text 
                        its = i.font.italic
                        if its:
                            text = f"<i>{text}<\i>"
                        tp_text.append(text)

                '''Make flag to indicate whether sec 2 or 3 was filled in'''
                assert tp_text != []
                self.attribs["taxonomy_proposal"] = tp_text      
            except:
                self.attribs["taxonomy_proposal"] = self.parser_errors["taxonomy_proposal"] = self.parser_err_codes[0]

    def get_excel_name(self, _row, _cell_idx, row_idx, table):
        try:
            excel_fname = [i.text for i in table.rows[row_idx + 1].cells if not i.text.strip().replace(" ","") == ""]
            if excel_fname == []:
                return
            self.attribs["excel_fname"] = excel_fname
        except:
            self.attribs["excel_fname"] = self.parser_errors["excel_fname"] = self.parser_err_codes[1]

    def get_proposed_changes(self, _row, _cell_idx, row_idx, table):
        if self.do_optional:
            proposals = []
            counter = 1
            while True:
                try:
                    row_text = [i.text for i in table.rows[row_idx + counter].cells]
                except IndexError:
                    '''End of table'''
                    break
                counter += 1
                if row_text[1] != "":
                    proposals.append(row_text[0])
            self.attribs["Proposed_taxonomic_changes"] = proposals
    
    def get_vanity_names(self, _row, _cell_idx, row_idx, table):
        if self.do_optional:
            vanity_names = []
            counter = 2
            while True:
                try:
                    row_text = [i.text for i in table.rows[row_idx + counter].cells]
                except IndexError:
                    '''End of table'''
                    break
                counter += 1
                if row_text[1] != "":
                    vanity_names.append(row_text)
            self.attribs["Taxon_vanity_names"] = vanity_names

    def main(self) -> str:
        '''Iterate over each table element, call parser functions, save.'''
        document = Document(f"{self.in_dir}{self.fname}")
        for table in document.tables:
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    for para in cell.paragraphs:
                        para_header = "".join([i.text for i in para.runs]).strip(" ")
                        if para_header in self.section_fns.keys():
                            self.section_fns[para_header](row, cell_idx, row_idx, table)
                        if para_header == "Abstract":
                            '''Increment index for measuring which section's abstract is being parsed'''
                            self.which_section = 3
        err_fname = self.collate_errors()
        self.save_json()
        self.make_summary()
        return err_fname
    
if __name__=="__main__":
    '''Input args'''
    in_dir = "data/"
    out_dir = "output/"
    do_optional = True
    error_logs = []
    if not os.path.exists(out_dir):
        os.mkdir(out_dir)

    '''Run parser'''
    for file in os.listdir(in_dir):
        clf = Parse(f"{in_dir}{file}")
        code, data, errors = clf.main()
        all_data[code] = data
        error_logs.append(errors)

    save_json(all_data, fname)

if __name__ == "__main__":
    fname = "plant_virus"
    main(fname)

