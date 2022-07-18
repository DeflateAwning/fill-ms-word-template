import sys
import os
import comtypes.client

import easygui as g
import glob
from tqdm import tqdm

from loguru import logger
logger.add(sys.stderr, format="{time} {level} {message}", filter=__file__, level="INFO")


wdFormatPDF = 17

def convert_docx_to_pdf(in_file, out_file):
    in_file = os.path.abspath(in_file)
    out_file = os.path.abspath(out_file)

    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

def do_it():
    folder_path = g.diropenbox("Select a dir with .docx files to make into .pdf files.")
    if folder_path is None:
        logger.error('No folder selected. Stop.')
        return

    for file_path in tqdm(glob.glob(os.path.join(folder_path, "*.docx"))):
        out_file = file_path.replace('.docx', '.pdf')
        logger.info(f"Start Filename: '{file_path}'.")
        logger.info(f"Dest. Filename: '{out_file}'.")

        convert_docx_to_pdf(file_path, out_file)

        logger.info(f"Done.")

    logger.info("All done.")

if __name__ == "__main__":
    do_it()
