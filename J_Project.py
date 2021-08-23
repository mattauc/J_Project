import extract
import manage_excel
import openpyxl
import json

def locate_name_date(data):

    report_info = []

    for table in data[-1]:
        #Grabbing report name, converting information to list
        name_row = table.iloc[2]
        name_list = name_row.values.tolist()

        #Grabbing report date
        date_row = table.iloc[7]
        date_list = date_row.values.tolist()

        report_info.append(name_list[0].replace("Norris, Jacqui ",""))
        #Modifying date to suit excel conditions
        date_list = date_list[0].split(" ")
        report_info.append(" ".join(date_list[2:6]))
        break

    return report_info

def parse_values(row, data_values):
    
    #Modifying strip of data, removing bad values and creating list
    data = str(row).replace("- Total", "(Total)")
    data = str(data).replace("+", "").replace("-", "").replace("    ", "  ")
    data = data.split("  ")

    profile = open('config.json')
    profile_data = json.load(profile)

    #Appending value to key
    if len(data) > 1 and data[0] in data_values:

        if len(profile_data['profiles'][data[0]]) == 0:
            data_values[data[0]] = " ".join(data[1:])
        else:
            data_values[data[0]] = data[1].split(" ")[0]
    return

def retrieve_data(data):

    collect = False
    #Initializes dictionary for data extraction
    data_values = extract.row_profile()

    for num, table in enumerate(data[-1]):
        count = 9
        while count < len(table)-4:

            #Converts row dataFrame into list and parses values
            row = table.iloc[count].values.tolist()
            parse_values(row[0], data_values)
         
            count += 1
    return data_values

def main():

    filenames = extract.find_pdfs() 
    #Grabs workbook
    workbook = manage_excel.create_workbook()
    
    for name in filenames:
        
        #Extract relevant data
        data = extract.pdf_to_text(name)
        report_info = locate_name_date(data)
        data_values = retrieve_data(data)

        #Apply data to workbook
        worksheet = manage_excel.create_sheet(report_info, data_values, workbook)
        manage_excel.update_column(report_info, data_values, worksheet)

    #Saving and closing workbook
    workbook.save(r'Remdesvir_trial.xlsx')
    workbook.close()

    print("Complete!")
    return

main()