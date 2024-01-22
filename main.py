import os
import openpyxl

def scan_for_initial_assesment(batch):

    #create the path from the batch given
    path='T:\data\R/' + batch

    #scan the directory and locate the .xlsx file
    for filename in os.listdir(path):
        if '.xlsx' and 'template' in filename: #make sure you always use a template excel file that NEVER changes name
            optical_yield_template=filename

    #join together to get a full directory string
    full_directory=os.path.join(path,optical_yield_template)

    #Open the Excel file
    workbook=openpyxl.load_workbook(full_directory)
    #reading the data from the first sheet
    sheet=workbook.active

    completed_panels = []
    non_completed_panels = []
    # Access and scan through the data in the Excel file
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[4]==None:
            #make sure the values in the list are unique
            if row[2] not in non_completed_panels:
                non_completed_panels.append(row[2])
        else:
            #make sure the values in the list are unique
            if row[2] not in completed_panels:
                completed_panels.append(row[2])

    print(non_completed_panels)
    print(completed_panels)



        #if row(3)=='None':
         #   print('Panel is not done:',row(2))



    #go to the directory given the batch ID input: R423
    #find the optical_yield_template.xlsx file
    #acess it and open it
    #scan to see what is populated





# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    scan_for_initial_assesment('R423')


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
