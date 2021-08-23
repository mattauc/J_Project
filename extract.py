from tabula import read_pdf
import json
import os

def check_collected(filename):

    #Reads the txt file containing all sorted pdfs
    try:
        collected_data = open("completion.log", 'r')
    except FileNotFoundError:
        return True

    lines = collected_data.readlines()
    for line in lines:

        if (line.strip() == filename):
            collected_data.close()
            return False

    collected_data.close()
    return True

def append_filename(filename):

    #Appends a filename to the collected txt
    try:
        collected_data = open("completion.log", 'a')
    except FileNotFoundError:
        collected_data = open("completion.log", 'w')

    collected_data.write(filename + "\n")
    collected_data.close()
    return

def find_pdfs():

    filenames = []

    #Iterates through directory, locating usable pdfs
    for filename in sorted(os.listdir("PDFdata/")):
        if filename.endswith(".PDF") or filename.endswith(".pdf") and check_collected(filename):
            #Writes to collected txt
            append_filename(filename)

            filenames.append(filename)
    return filenames

def pdf_to_text(pdf_file):

    data = []
    #Converts pdf into a dataFrame for ease of access
    dfs = read_pdf("PDFdata/" + pdf_file, pages='all', multiple_tables=True, guess=False)
    data.append(dfs)
    
    return data

def row_profile():

    #Reads profile data to create dictionary
    data_values = {}    
    
    profile = open('config.json')
    data = json.load(profile)

    for key in data['profiles']:
        data_values[key] = None

    profile.close()
    return data_values