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
        self.raw_data = raw_contents

    def get_metadata(self):
        '''Get id code and title'''
        datum = self.raw_data[0]
        self.attribs["code"] = datum[[i for i, x in enumerate(datum) if "Code assigned:" in x][0] + 1]
        if self.attribs["code"] == "": breakpoint()
        self.attribs["title"] = datum[[i for i, x in enumerate(datum) if "Short title:" in x][0]].replace("Short title: ", "")
        self.attribs["study_grp"] = self.raw_data[4][0]

    def get_authors(self):
        authors = self.raw_data[1]
        addresses = self.raw_data[2]
        if authors == "": breakpoint()
        self.attribs["authors"] = {}
        self.attribs["authors"]["names"] = [i.strip().replace(".", "") for i in authors[0].replace(";", ",").split(",") if not i.strip().replace(" ", "") == ""]
        self.attribs["authors"]["emails"] = [i.strip() for i in authors[1].replace(";", ",").split(",") if not i.strip().replace(" ", "") == ""]
        self.attribs["authors"]["addresses"] = [f"{i}" for i in addresses[0].replace("]",")").replace("[","").replace(")",")@~").split("@~") if not i.strip().replace(" ", "") == ""]
        self.attribs["authors"]["corr_author"] = self.raw_data[3][0]

    def get_content_sub13(self):
        self.attribs["submission_date"] = self.raw_data[6][1]
        try:
            self.attribs["revision_date"] = self.raw_data[6][3]
        except:
            self.attribs["revision_date"] = ""
        if ".xl" in self.raw_data[8][0]:
            excel_box = 8
        elif ".xl" in self.raw_data[9][0]:
            excel_box = 9
        elif ".xl" in self.raw_data[10][0]:
            excel_box = 10
        elif ".xl" in self.raw_data[11][0]:
            excel_box = 11
        else: breakpoint()

        self.attribs["excel_fname"] = self.raw_data[excel_box][0]
        self.attribs["abstract"] = self.raw_data[excel_box+1][0]
        self.attribs["proposal_text"] = " ".join(self.raw_data[excel_box+2])

    def get_content_15(self):
        if self.raw_data[10][0] == "": # HACKY AF :(
            self.attribs["submission_date"] = self.raw_data[9][1]
            try:
                self.attribs["revision_date"] = self.raw_data[9][3]
            except:
                self.attribs["revision_date"] = ""
            self.attribs["excel_fname"] = self.raw_data[11][0]
            self.attribs["abstract"] = self.raw_data[12][0]
            self.attribs["proposal_text"] = " ".join(self.raw_data[13])
        elif ".xl" in self.raw_data[9][0]:
            try:
                self.attribs["submission_date"] = self.raw_data[8][1]
            except: breakpoint()
            try:
                self.attribs["revision_date"] = self.raw_data[8][3]
            except:
                self.attribs["revision_date"] = ""
            self.attribs["excel_fname"] = self.raw_data[9][0]
            self.attribs["abstract"] = self.raw_data[10][0]
            self.attribs["proposal_text"] = " ".join(self.raw_data[11])

        else:
            self.attribs["submission_date"] = self.raw_data[9][1]
            try:
                self.attribs["revision_date"] = self.raw_data[9][3]
            except:
                self.attribs["revision_date"] = ""
            self.attribs["excel_fname"] = self.raw_data[10][0]
            self.attribs["abstract"] = self.raw_data[11][0]
            self.attribs["proposal_text"] = " ".join(self.raw_data[12])

    def get_content_13(self):
        self.attribs["submission_date"] = self.raw_data[6][1]
        try:
            self.attribs["revision_date"] = self.raw_data[6][3]
        except:
            self.attribs["revision_date"] = ""
        if ".xl" in self.raw_data[8][0]:
            excel_box = 8
        elif ".xl" in self.raw_data[9][0]:
            excel_box = 9
        elif ".xl" in self.raw_data[10][0]:
            excel_box = 10
        elif ".xl" in self.raw_data[11][0]:
            excel_box = 11
        else: breakpoint()

        self.attribs["excel_fname"] = self.raw_data[excel_box][0]
        self.attribs["abstract"] = self.raw_data[excel_box+1][0]
        self.attribs["proposal_text"] = " ".join(self.raw_data[excel_box+2])


    def get_content_14(self):
        try:
            self.attribs["submission_date"] = self.raw_data[9][1]
        except: 
            self.attribs["submission_date"] = self.raw_data[9][0]
        try:
            self.attribs["revision_date"] = self.raw_data[9][3]
        except:
            self.attribs["revision_date"] = ""

        if ".xl" in self.raw_data[8][0]:
            excel_box = 8
        elif ".xl" in self.raw_data[9][0]:
            excel_box = 9
        elif ".xl" in self.raw_data[10][0]:
            excel_box = 10
        elif ".xl" in self.raw_data[11][0]:
            excel_box = 11
        else: breakpoint()

        self.attribs["excel_fname"] = self.raw_data[excel_box][0]
        self.attribs["abstract"] = self.raw_data[excel_box+1][0]
        self.attribs["proposal_text"] = " ".join(self.raw_data[excel_box+2])

    def how_many_lengths_are_there_guys_seriously(self):
        try:
            self.attribs["submission_date"] = self.raw_data[9][1]
        except: 
            self.attribs["submission_date"] = self.raw_data[9][0]
        try:
            self.attribs["revision_date"] = self.raw_data[9][3]
        except:
            self.attribs["revision_date"] = ""
    
        if ".xl" in self.raw_data[8][0]:
            excel_box = 8
        elif ".xl" in self.raw_data[9][0]:
            excel_box = 9
        elif ".xl" in self.raw_data[10][0]:
            excel_box = 10
        elif ".xl" in self.raw_data[11][0]:
            excel_box = 11
        elif ".xl" in self.raw_data[12][0]:
            excel_box = 12
        else: breakpoint()

        self.attribs["excel_fname"] = self.raw_data[excel_box][0]
        self.attribs["abstract"] = self.raw_data[excel_box+1][0]
        self.attribs["proposal_text"] = " ".join(self.raw_data[excel_box+2])


    def get_content_16(self):
        try:
            self.attribs["submission_date"] = self.raw_data[9][1]
        except: 
            self.attribs["submission_date"] = self.raw_data[9][0]
        try:
            self.attribs["revision_date"] = self.raw_data[9][3]
        except:
            self.attribs["revision_date"] = ""

        if ".xl" in self.raw_data[8][0]:
            excel_box = 8
        elif ".xl" in self.raw_data[9][0]:
            excel_box = 9
        elif ".xl" in self.raw_data[10][0]:
            excel_box = 10
        elif ".xl" in self.raw_data[11][0]:
            excel_box = 11
        elif ".xl" in self.raw_data[12][0]:
            excel_box = 12
        else: breakpoint()

        self.attribs["excel_fname"] = self.raw_data[excel_box][0]
        self.attribs["abstract"] = self.raw_data[excel_box+1][0]
        self.attribs["proposal_text"] = " ".join(self.raw_data[excel_box+2])


    def main(self):
        self.scrape()
        self.get_metadata()
        self.get_authors()

        if len(self.raw_data) < 13:
            '''Short form format'''
            self.get_content_sub13()
        elif len(self.raw_data) == 13:
            self.get_content_13()
        elif len(self.raw_data) == 14:
            self.get_content_14()
        elif len(self.raw_data) == 15:
            self.get_content_15()
        elif len(self.raw_data) == 16:
            self.get_content_16()
        else: 
            self.how_many_lengths_are_there_guys_seriously()

        # if "2023.072B" in self.attribs["code"]: breakpoint()

        print(f"TEXT PARSED: {self.attribs['code']}")
        return self.attribs["code"], self.attribs, "" # TODO errors

def save_json(data, fname) -> None:
    '''Dump results to machine-readable format'''
    with open(f"{fname}.json", "w") as outfile: 
        json.dump(data, outfile)                    

if __name__ == "__main__":
    in_dir = "data/"
    out_dir = "output/"
    do_optional = False
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

    save_json(all_data, "animal_+ssrna")

