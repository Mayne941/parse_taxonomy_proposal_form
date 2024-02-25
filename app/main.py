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
        self.attribs["title"] = datum[[i for i, x in enumerate(datum) if "Short title:" in x][0]].replace("Short title: ", "")
        self.attribs["study_grp"] = self.raw_data[4][0]
        self.attribs["submission_date"] = self.raw_data[6][1]
        try:
            self.attribs["revision_date"] = self.raw_data[6][3]
        except:
            self.attribs["revision_date"] = ""
        self.attribs["excel_fname"] = self.raw_data[8][0]

    def get_authors(self):
        authors = self.raw_data[1]
        addresses = self.raw_data[2]
        self.attribs["authors"] = {}
        self.attribs["authors"]["names"] = [i.strip().replace(".", "") for i in authors[0].replace(";", ",").split(",")]
        self.attribs["authors"]["emails"] = [i.strip() for i in authors[1].replace(";", ",").split(",")]
        self.attribs["authors"]["addresses"] = [f"{i}" for i in addresses[0].replace("]",")").replace("[","").replace(")",")@~").split("@~") if not i == ""]
        self.attribs["authors"]["corr_author"] = self.raw_data[3][0]

    def get_content(self):
        self.attribs["sg_comments"] = self.raw_data[5][0]
        self.attribs["ec_comments"] = self.raw_data[7][0] if not self.raw_data[7][0] == "Is any taxon name used here derived from that of a living person (Y/N)" else ""
        self.attribs["abstract"] = self.raw_data[9][0]
        self.attribs["proposal_text"] = self.raw_data[10][0]

    def main(self):
        self.scrape()
        self.get_metadata()
        self.get_authors()
        self.get_content()

        return self.attribs["code"], self.attribs, "" # TODO errors

def save_json(data) -> None:
    '''Dump results to machine-readable format'''
    with open(f"test.json", "w") as outfile: 
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

    save_json(all_data)

