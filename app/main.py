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

        more_content = soup.find_all('p')
        raw_contents = []
        for tag in more_content:
            table = tag.find_next_sibling('w:tbl')
            table_contents = []
            if table:
                for wtc in table.findChildren('w:tc'):
                    cell_text = ''
                    for wr in wtc.findChildren('w:r'):
                        if "<w:i/>" in str(wr) and not wr.text == " ":
                            cell_text += f"<i>{wr.text}</i>"
                        elif "w:fldChar" in str(wr): 
                            cell_text = "<Table data voided by parser>"
                        else:
                            cell_text += wr.text
                            
                    table_contents.append(cell_text.replace("</i><i>",""))
            if not table_contents == [] and not table_contents in raw_contents:
                raw_contents.append(table_contents)

        for doc in raw_contents:
            '''Fix excel path'''
            for item in doc:
                item = item.replace(".xlxs", ".xlsx")

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

        '''Count length of boxes; guess where submission and excel boxes are, pass to scraper'''
        if len(self.raw_data) <= 13:
            self.get_content(6,8)
        elif len(self.raw_data) == 14:
            self.get_content(9,8)
        elif len(self.raw_data) == 15:
            self.get_content(10,8)
        else:
            self.get_content(9,8)

        # if self.attribs["code"] == "2023.005P": breakpoint()

        print(f"TEXT PARSED: {self.attribs['code']}")
        return self.attribs["code"], self.attribs, "" # TODO errors

def save_json(data, fname) -> None:
    '''Dump results to machine-readable format'''
    with open(f"{fname}.json", "w") as outfile: 
        json.dump(data, outfile)                    

def main(fname):
    in_dir = "data/"
    out_dir = "output/"
    all_data = {}
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

