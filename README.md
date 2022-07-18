# fill-ms-word-template
A tool to fill multiple Microsoft Office Word templates based on an Excel table, and then save them to PDFs to print easily.

## Tools Summary
There are two separate tools. These tools may often be used in this order:

* `fill_word_doc.py`: Based on a template Word document and an Excel table containing a list of replacements, generate multiple copies of the Word document template, filled with data from the table.
* `export_all_word_to_pdf.py`: Export a folder containing multiple Word documents to PDFs (each word file in a separate PDF).

## Usage Instructions
Install dependencies with `python3 -m pip install -r requirements.txt`.

### `fill_word_doc.py`: Fill a Word document template from a table.
Refer to the sample input-output sample in `/sample (fill_word_doc)/`.

1. Create a Word document to use as a template to be filled. The template can contain either placeholder fields (like `#name_of_item#`), or contain the filled values for the first row (like if the name of the item was `stapler`, then the template can be pre-filled with `stapler`).
2. Create an Excel spreadsheet. The names of the columns do not matter (i.e., they are ignored). The first row should contain the values from the Word document (like `#name_of_item#` or `stapler`).
3. Run the script (double-click it on Windows, or use `python3 fill_word_doc.py` from a Terminal).
4. When file dialog boxes pop up, follow the instructions in the title bar of the pop up windows.
    1. First dialog: Select the Word document template.
    2. Second dialog: Select the Excel table.
    3. Third dialog: Select the output folder.
5. The script will continue running, and will generate files (one file per table row).

* Any columns in the table file with column names that start with an underscore are ignored completely.
* The replacements are performed on the filename as well.

### `export_all_word_to_pdf.py`: Export a folder containing multiple Word documents to PDFs.
1. Run the script (double-click it on Windows, or use `python3 export_all_word_to_pdf.py` from a Terminal).
2. When file dialog box pops up, select the folder containing multiple Word document files.
3. The script will run, and will output the documents as PDF files in the same folder as the Word documents.

## Attribution
* This project is released under the MIT License, so you can do anything you want with it!
* Please star this repo if it is useful to you.
* Contributions are welcome!
