import os
import pyttsx

from docx import Document


def read_table(doc_file_name):
    file_path = os.getcwd() + "/" + doc_file_name
    try:
        row_id = 0
        word_doc = Document(file_path)
        engine = pyttsx.init()
        for table in word_doc.tables:
            for row in table.rows:
                row_id += 1
                column_id = 0
                for cell in row.cells:
                    column_id += 1
                    print '{} {} {} {}: {}'.format("row", row_id, "column", column_id, cell.text)
                    engine.say(cell.text)
                    engine.runAndWait()
    except Exception, e:
        print '{} {}'.format("Error occurred due to ", str(e))


print ("\nSteps to run the program")
print ("------------------------")
print ("Step1: Copy and paste the word document file in the same location where this python file residing.")
print ("Step2: Provide the name of the document file as <filename>.docx\n")

file_name = raw_input("Name of the document file: ")

if (file_name.split(".")[-1]) == "docx":
    read_table(file_name)
else:
    print "Invalid file format!"
