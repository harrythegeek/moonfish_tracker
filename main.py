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
        if row[4] is None:
            #make sure the values in the list are unique
            if row[2] not in non_completed_panels:
                non_completed_panels.append(row[2])
        else:
            #make sure the values in the list are unique
            if row[2] not in completed_panels:
                completed_panels.append(row[2])

    return completed_panels, non_completed_panels

def scan_images_directory(batch):

    completed_panels = []
    non_completed_panels = []

    path = 'T:\data\R/' + batch + '/optical/images'

    for root, dirs, files in os.walk(path):
        for name in files:

            #look for the panel number without the 0 contained in the filename (example: panel 3 not panel 03)
            #SOP make sure we don't get confused using panel 0X instead of panel X
            if name[6]=='1' or name[6]=='2' or name[6]=='3' or name[6]=='4' or name[6]=='5' or name[6]=='6':
                if name[6] not in completed_panels:
                    completed_panels.append(name[6])
            else:
                continue

    return completed_panels

def scan_resistance_check(batch):

    #create the path from the batch given
    path='T:\data\R/' + batch + '/electrical'

    #scan the directory and locate the .xlsx file
    for filename in os.listdir(path):
        if '.xlsx' and 'resistance_checks' in filename: #make sure you always use a template excel file that NEVER changes name
            resistance_checks=filename

    #join together to get a full directory string
    full_directory=os.path.join(path,resistance_checks)

    #Open the Excel file
    workbook=openpyxl.load_workbook(full_directory)
    #reading the data from the first sheet
    sheet=workbook.active


    panels_done=[]
    cells_done=[]
    # Access and scan through the data in the Excel file
    for row in sheet.iter_rows(min_row=2, values_only=True):
        print(row)
        if row[3] and row[4] and row[5] is not None:
            panel_ID=row[1]
            cell_ID=row[2]

            panels_done.append(panel_ID)
            cells_done.append(cell_ID)


 #Press the green button in the gutter to run the script.
if __name__ == '__main__':

    #first step
    completed_initial_assesssment,not_completed_initial_assesssment=scan_for_initial_assesment('R423')
    #second step
    completed_global_motherglass_images=scan_images_directory('R423')
    #third step
    scan_resistance_check('R423')

    print('First step: Initial Assessment')
    print('Panels that have been through initial assessment',completed_initial_assesssment)
    print('Panels that have not been through initial assessment',not_completed_initial_assesssment)
    print()
    print('Second step: Global mother glass images')
    print('Panels that have global images taken',completed_global_motherglass_images)


