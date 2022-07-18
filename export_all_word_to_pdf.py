#!/usr/bin/env python3

# MIT License

# Copyright (c) 2022 DeflateAwning

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

###############################################################################

# Purpose: Export a folder containing multiple Word documents to PDFs (each word file in a separate PDF).

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
