from docx import Document
import numpy as np
import json
import os


class Strip:
    def __init__(self, fname, in_dir, out_dir) -> None:
        self.fname = fname
        self.in_dir = in_dir
        self.out_dir = out_dir
        self.attribs = {}
        self.which_section = 2
        self.section_fns = {
            "Title": self.populate_title,
            "Code assigned:": self.populate_id,
            "Author(s), affiliation and email address(es) –": self.populate_authors,
            "Corresponding author(s)": self.get_main_author,
            "Sub-committee": self.get_subcommittee,
            "List the ICTV Study Group(s) that have seen or who have involved in creating this proposal.": self.get_study_groups,
            "Submission date": self.get_subm_date,
            # "Optional – complete if formally voted on by an ICTV Study Group": self.populate_group_vote,
            # "Executive Committee Meeting Decision code": self.get_meeting_decision,
            # "Comments from the Executive Committee": self.get_comments,
            # "Response of proposer": self.get_response,
            "Revision date": self.get_rev_date,
            "Abstract": self.get_abstract
        }

    def populate_title(self, row, cell_idx, *_):
        '''Populate title'''
        title_text = [i.text for i in row.cells[cell_idx+1].paragraphs[0].runs]
        title_it_mask = [i.font.italic for i in row.cells[cell_idx+1].paragraphs[0].runs]
        it_indices = np.argwhere(np.array(title_it_mask) != None)
        for i in it_indices:
            title_text[i[0]] = f"<i>{title_text[i[0]]}<\i>"
        self.attribs["Title"] = title_text

    def populate_id(self, row, cell_idx, *_):
        self.attribs["Id_code"] = [i.text for i in row.cells[cell_idx+1].paragraphs[0].runs]

    def populate_authors(self, row, cell_idx, row_idx, table):
        authors = []
        counter = 2
        while True:
            try:
                row_text = [i.text for i in table.rows[row_idx + counter].cells]
            except IndexError:
                '''End of table'''
                break
            counter += 1
            if row_text == ["", "", ""]:  
                '''Blank line'''
                continue
            else:
                '''Author details'''
                authors += row_text
        self.attribs["Authors"] = authors

    def get_main_author(self, row, cell_idx, row_idx, table):
        self.attribs["Corr_author"] = [i.text for i in table.rows[row_idx + 1].cells]

    def get_subcommittee(self, row, cell_idx, row_idx, table):
        counter = 0
        subcommittees = []
        while True:
            try:
                row_text = [i.text for i in table.rows[row_idx + counter].cells]
            except IndexError:
                '''End of table'''
                break
            counter += 1
            if row_text == ['Sub-committee', 'X', 'Sub-committee', 'X']:
                '''Ignore headers'''
                continue
            if row_text[1] != "":
                '''Left column match'''
                subcommittees.append(row_text[0])
            if row_text[3] != "":
                '''Right column match'''
                subcommittees.append(row_text[2])
        self.attribs["Subcommittees"] = subcommittees

    def get_study_groups(self, row, cell_idx, row_idx, table):
        self.attribs["Study_groups"] = [i.text for i in table.rows[row_idx + 1].cells]

    def get_subm_date(self, row, cell_idx, *_):
        [i.text for i in row.cells[cell_idx+1].paragraphs[0].runs]

    def populate_group_vote(self, row, cell_idx, row_idx, table):
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

    def get_meeting_decision(self, row, cell_idx, row_idx, table):
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

    def get_comments(self, row, cell_idx, row_idx, table):
        comments = []
        counter = 1
        while True:
            try:
                row_text = [i.text for i in table.rows[row_idx + counter].cells]
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

    def get_response(self, row, cell_idx, row_idx, table):
        response = []
        counter = 1
        while True:
            try:
                row_text = [i.text for i in table.rows[row_idx + counter].cells]
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

    def get_rev_date(self, row, cell_idx, row_idx, table):
        self.attribs["Revision_date"] = [i.text for i in row.cells[cell_idx+1].paragraphs[0].runs]

    def get_abstract(self, row, cell_idx, row_idx, table):
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

        if self.which_section == 2:
            self.attribs["Tp_type"] = ["Non-taxonomic proposal"]

        elif self.which_section == 3:
            assert not "Tp_type" in self.attribs.keys(), "Error: User has filled in both section 2 + 3"
            self.attribs["Tp_type"] = ["Taxonomic proposal"]
        self.attribs["Tp_abstract"] = abstract        

    def make_summary(self):
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

    def main(self):
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

        with open(f"{self.out_dir}{self.attribs['Id_code'][0]}.json", "w") as outfile: 
            json.dump(self.attribs, outfile)                

        self.make_summary()

if __name__=="__main__":
    in_dir = "data/"
    out_dir = "output/"
    if not os.path.exists(out_dir):
        os.mkdir(out_dir)
    for file in os.listdir(in_dir):
        strip = Strip(file, in_dir, out_dir) 
        strip.main()