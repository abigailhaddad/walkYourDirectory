# -*- encoding: utf-8 -*-
# os, time, and pandas are packages everyone is going to have
# some of these are less so, but I was able to install all of them on my gfe
# if you are not able to, you can comment them out and run this with text_pull=False

import os
import pandas as pd
import time
from openpyxl import load_workbook
from docx import Document
import pptx
import PyPDF2


def conditions(file):
    """
    :param file: the name of the doc_file
    :return: True if it doesn't have $ and is isn't thumbs.db, false otherwise
    """
    if "$" in file:
        return False
    if "Thumbs.db" in file:
        return False
    else:
        return True


def getcsvtext(csv_file):
    """
    :param csv_file: filename of CSV doc_file
    :return: all of the text in that doc_file
    """
    data = pd.read_csv(csv_file).dropna(how='all').dropna(axis='columns', how='all')
    text = data.to_string().strip().replace("  ", " ").replace("\n", "")
    return text


def runthroughpulls(file):
    """
    :param file: doc_file name
    :return: doc_file text if able to read it, "" otherwise
    """
    lower = file.lower()
    text = ""
    try:
        if ".xls" in lower:
            creator, modified, text = scrape_excel(file)
        elif ".ppt_file" in lower:
            creator, modified, text = scrape_ppt(file)
        elif ".pdf" in lower:
            creator, modified, text = scrape_pdf(file)
        elif ".doc" in lower:
            creator, modified, text = scrape_word(file)
        elif ".csv" in lower:
            creator, modified = ("", "")
            text = getcsvtext(file)
        else:
            creator, modified, text = "", "", ""
    except:
        creator, modified, text = "", "", ""
    return creator, modified, text


def gettime(file):
    """
    :param file: name of doc_file
    :return: creation and modify time of that doc_file, if able to pull them
    """
    try:
        creation_time = time.ctime(os.path.getctime(file))
        modify_time = time.ctime(os.path.getmtime(file))
        return creation_time, modify_time
    except:
        return "", ""


def runwords(text, keywords):
    """
    :param text: the text that you want to search
    :param keywords: the keywoards that you want to find
    :return: The subset of the keywords that were found.
    """
    keywords_found = []
    for substring in keywords:
        if substring in text:
            keywords_found.append(substring)
    return ",".join(keywords_found)


def get_file_list(folder, keywords, text_pull):
    """
    :param folder: The name of the directory.
    :param keywords: a list of keywords to be searched.
    :param text_pull: a boolean value that
    :return: a data frame containing metadata on each xlsx doc_file and text data from that doc_file
    """
    all_files = []
    for dirpath, dirnames, filenames in os.walk(folder):
        for filename in filenames:
            file = os.path.join(dirpath, filename)
            if conditions(file):
                if text_pull:
                    creator, modified, text = runthroughpulls(file)
                else:
                    creator, modified, text = "", "", ""
                creation_time, modify_time = gettime(file)
                extension = file.split(".")[-1]
                row = dirpath, filename, text, extension, creator, modified, creation_time, modify_time
                all_files.append(row)
    df = pd.DataFrame(all_files)
    df.columns = ["location", "doc_file name", "doc_file text", "extension", "creator", "modified by", "creation time",
                  "modified time"]
    df['Found keyword'] = df['doc_file text'].astype(str).str.lower().astype(str).apply(runwords, args=keywords)
    df = df.drop_duplicates()
    if not text_pull:
        df = df.drop(columns=["creator", "modified by", "doc_file text", "Found keyword"])
    return df


def word_list():
    """
    :return: a list of words we're going to look for in our text
    """
    keywords = ["model", "inputs", "assumptions", "outputs", "actual", "predicted", "attributes"]
    return keywords


def scrape_word(doc_file):
    """
    :param doc_file: csv_file name (word doc)
    :return: creator, modified by, text of csv_file
    """
    try:
        document = Document(doc_file)
        creator, modified_by = get_file_info(document)
        text = get_text_docx(document)
        return creator, modified_by, text
    except:
        return "", "", ""


def get_file_info(content):
    """
    :param content: content of a document or ppt_file
    :return: creator, modified by
    """
    try:
        creator = content.core_properties.author
        modified_by = content.core_properties.last_modified_by
        return creator, modified_by
    except:
        return "", ""


def scrape_ppt(ppt_file):
    """
    :param ppt_file: filename of a ppt_file doc_file
    :return: creator, modified by, filetext
    """
    try:
        ppt = pptx.Presentation(ppt_file)
        creator, modified_by = get_file_info(ppt)
        text = get_text_ppt(ppt)
        return creator, modified_by, text
    except:
        return "", "", ""


def scrape_excel(excel_file):
    """
    :param excel_file: the filename of the excel pdf_file.
    :return: creator, modified by, text of doc_file
    """
    try:
        wb = load_workbook(excel_file)
        creator = wb.properties.creator
        modified_by = wb.properties.lastModifiedBy
        text = get_text_excel(wb)
        return creator, modified_by, text
    except:
        return "", "", ""


def get_text_excel(wb):
    """
    :param wb: an excel workbook
    :return: string with text containing all the text from the workbook
    """
    cell_rows = []
    for sheet in wb.worksheets:
        for row_cells in sheet.iter_rows():
            cells = [cell.value for cell in row_cells if cell.value is not None]
            cell_rows.append(cells)
    flat_list = [str(item) for sublist in cell_rows for item in sublist]
    return " ".join(flat_list)


def scrape_pdf(pdf_file):
    """
    :param pdf_file: filename of PDF
    :return: text from PDF with the caveat that PDFs are twitchy and might not get everything
    """
    pdf_file_object = open(pdf_file, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(pdf_file_object)
    creator = pdf_reader.getDocumentInfo().author
    pages = pdf_reader.numPages
    string = ""
    for num in range(0, pages):
        page_object = pdf_reader.getPage(num)
        string = string + page_object.extractText()
    pdf_file_object.close()
    text = string.replace("\n", "")
    return creator, "", text


def get_text_ppt(ppt_file):
    """
    :param ppt_file: filename of PPT
    :return: filename of PPT
    """
    text = ""
    for slide in ppt_file.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text = text + shape.text + " "
    return text


def get_text_docx(document):
    """
    :param document: a document in docx
    :return: the text from that document
    """
    all_text = ""
    for para in document.paragraphs:
        all_text = all_text + " " + para.text.strip()
    return all_text.strip().replace("  ", " ")


def get_formulas(string, extension):
    """
    :param string: string of text
    :param extension: file extension
    :return: if this is an xlsx, this will return a list of the formulas in the filetext
    """
    if extension == "xlsx":
        the_list = string.split(" ")
        formulas = [i for i in the_list if i.startswith("=")]
        return ",".join(formulas)
    else:
        return ""


def main(folder, keywords=word_list(), textpull=True, formulas=False):
    """
    :param folder: The directory we want to search.
    :param keywords: an optional list of keywords ex. ['dog', 'cat']
    :param textpull: param for what text you want to pull out of documents.
    :param formulas: param to decide if you want to pull formulas out of excel files.
    :return: a dataframe with text data and metadata from the excel files there
    """
    df = get_file_list(folder, keywords, textpull)
    if formulas:
        df['formulas'] = df.apply(lambda x: get_formulas(x['doc_file text'], x['extension']), axis=1)
    return df


if __name__ == "__main__":
    path = input("Paste the directory ou want to walk\n")
    main(path)