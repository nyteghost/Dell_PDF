import io
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.layout import LAParams
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
import pandas as pd
import re
import os

filepath = ""
fileExt = ".pdf"


def pdf_to_text(path):
    """
    Convert a pdf file to a string.
    """
    with open(path, 'rb') as fp:
        rsrcmgr = PDFResourceManager()
        outfp = io.StringIO()
        laparams = LAParams()
        device = TextConverter(rsrcmgr, outfp, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.get_pages(fp):
            interpreter.process_page(page)
    text = outfp.getvalue()
    return text


def stringToList(string):
    """
    Convert string to list
    """
    listRes = list(string.split(" "))
    return listRes


for file in os.listdir(filepath):
    if file.endswith(".pdf"):
        filename = file.replace('.pdf', '')
        output = filepath + filename + ".xlsx"

        mystring = pdf_to_text(filepath+file)
        after_mystring = mystring.partition("Tags")[2]  # get the string after the "Tags"
        pdfList = stringToList(after_mystring.replace('\n', ''))  # convert to list and remove newlines
        print(pdfList)

        col1 = "Service Tags"
        col2 = "Service Tags Fixed"
        col3 = "Service Tags Found"
        res = []
        for i in pdfList:
            removed_parenth = re.sub("\(.*?\)", "", i)
            s = re.sub(r'\W+', '', removed_parenth)
            res.append(s)
        df2 = pd.DataFrame({col1: pdfList, col2: res})
        print(df2)

        df3 = pd.DataFrame({col3: res})
        df3 = df3[df3[col3].str.strip().astype(bool)]
        df3 = df3.reset_index(drop=True)
        df3.index += 1
        print(df3)

        parenth = re.findall('\(.*?\)', after_mystring)
        value = None
        if parenth:
            for i in parenth:
                value = i.replace("(", "").replace(")", "")
            if value:
                if len(df3) == int(value):
                    print(f"Amount of tags found {len(df3)}, matches expected {value}.")
                    df3.to_excel(output, header=False, index=False)
                else:
                    print(f"Amount of tags expected{value}\nAmount of tags found{len(df3)}")
            else:
                print("No value found.")
        else:
            print("No parenthesis found.")

print()
