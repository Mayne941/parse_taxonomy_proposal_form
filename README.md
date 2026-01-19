# Parse ICTV Taxonomy Proposal Document
Python script to extract data from the ICTV taxonomy proposal form. Capable of batch processing and preserving formatting from original document. 

By @mayne941, @donaldsmithictv-cell & @psimmond, on behalf of @ICTV-Virus-Knowledgebase.

Outputs data to a machine-readable/database-compatible JSON document and a summary word document.

N.b. this code is in an alpha state. Expect bugs, breakpoints etc. Testing, documentation etc. will follow.

# Pre-requisites
1. Python 3.7 + 
1. ICTV taxonomy proposals: paired .docx and .xlsx documents, see https://ictv.global/files/proposals/pending 
1. Create csv document containing details of ICTV SC chairs, must have following columns: Subcommittee, Name, Affiliation, Email  (n.b. Subcommittee must match name in data and data folders exactly!)

Code expects you to have n folders in the base directory, containing two folders: data (paste .docx files here) and data_tables (paste .xlsx files here). Folder names are input in ```app.entrypoint```.  

# Usage
1. Pull repository.
1. Ensure pre-requisites in place (above).
1. Install Python libraries: ```pip install -r requirements.txt```
1. Run main script: ```python3 -m app.entrypoint```

