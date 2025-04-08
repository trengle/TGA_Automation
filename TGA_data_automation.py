import openpyxl
import os
import pandas as pd
import csv
# import matplotlib.pyplot as plt
from openpyxl.chart import LineChart, Reference
import glob
import create_demo_files
import sys

## ACTUAL PROGRAM STARTS ###
def main():
    '''
    I need a more user-friendly way to look for the files that need to be added to the master excel file.
    Perhaps glob module to find all csv files in the thumb drive folder?
    Or have the user manually input the file names?
    '''

    ###Glob.glob(\filepath\*.extension)
    ## def find_files_by_extension(directory, extension):
    ##     pattern = os.path.join(directory, f"*{extension}")
    ##     return glob.glob(pattern) 
    ## csv_files = find_files_by_extension(rf'{os.getcwd()}', ".csv")

    # Loops through CWD and searches for all csv files
    csv_files = glob.glob(rf"*.csv")
    if glob.glob(rf"*.csv"):
        print("CSV files found")
    else:
        print("No CSV files found")
        sys.exit(1)
    # Iterates through each of the files and process them
    for csv_file in csv_files:
        rowlist = []
        target1 = float(180)
        target2 = float(500)
        with open(f"{csv_file}") as file:
            reader = csv.DictReader(file)
            for row in reader:
                rowlist.append(row)
        closest_dict1 = min(rowlist, key=lambda x: abs(float(x['Temp']) - target1))
        closest_dict2 = min(rowlist, key=lambda x: abs(float(x['Temp']) - target2))
        # The 2 target Temp/Weight value pairs have been found, now will add those to the master>sample sheet, run the equation and add the result purity
        ## os.chdir('..')
        mb = openpyxl.load_workbook('master_wb.xlsx') 
        general_equation = mb['Equation']['A1'].value
        evaluated_equation = general_equation.replace("a", str(closest_dict1['Temp'])).replace("b", str(closest_dict1['Weight'])).replace("x", str(closest_dict2['Temp'])).replace("y", str(closest_dict2['Weight']))
        expression = evaluated_equation.split('=')[0] # Strip the "=z" part to isolate the expression
        # Use eval to calculate the value of the expression
        z = eval(expression)

    ### NEED TO COPY CSV DATA TO THIS FILE IN SHEET WITH SAMPLE NAME ###
        mb.create_sheet(csv_file.split('.')[0])
        sample_sheet = mb[rf"{csv_file.split('.')[0]}"]
        sample_sheet['A1'] = "Temp"
        sample_sheet['B1'] = "Weight"
        sample_sheet['C1'] = "Purity"
        sample_sheet['A2'] = closest_dict1['Temp']
        sample_sheet['B2'] = closest_dict1['Weight']
        sample_sheet['A3'] = closest_dict2['Temp']
        sample_sheet['B3'] = closest_dict2['Weight']
        sample_sheet['C2'] = z
        # Adjust column widths to fit their contents
        for column in sample_sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter  # Get column letter
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = max_length + 2  # Add padding
            sample_sheet.column_dimensions[column_letter].width = adjusted_width

        # def add_graph():
        #     # Extract data from specific columns
        #     x_data = []
        #     y_data = []
        #     for row in sample_sheet.iter_rows(min_row=2, max_row=sample_sheet.max_row, min_col=1, max_col=2, values_only=True):
        #         x_data.append(row[0])  # Assuming column 1 contains x-axis data
        #         y_data.append(row[1])  # Assuming column 2 contains y-axis data

        #     # Create the graph
        #     plt.plot(x_data, y_data, marker='o')
        #     plt.title('Graph Based on Excel Data')
        #     plt.xlabel('X-axis Label')  # Replace with your label
        #     plt.ylabel('Y-axis Label')  # Replace with your label
        #     plt.grid(True)

        #     # Show or save the graph
        #     plt.savefig('excel_graph.png')  # Save as an image file
        #     plt.show()  # Display the graph
        
        mb.save("master_wb.xlsx")
        print(rf"'{csv_file}' Data added to master excel file")

if __name__ == "__main__":
    main()
    # create_demo_files.create_demo_csv("random_sample001")
    # create_demo_files.create_demo_master_excel("Random_Master")
'''
    SAMPLE FORMAT(sample numb, date, size of sample)= Sample001 20250406 35mg
    this gets input into the TGA 
    (TGA is connected to a computer that run windows OS, but has its own display, which will say "done" or something)
    it will run experiment, it will take about 90 min
    put in thumbdrive
    there is an app on the windows computer that you export the csv data from app to thumbdrive

    take the thumbdrive to local computer
    turn csv into excel for staging
    NO DATA IS AUTOMATICALLY ENTERED INTO THE MASTER EXCEL FILE
    copy and paste the 2 k,v pairs into the master excel (labeled? which sheet? what is the current state of the master excel?)
    It is true that each sample number has its own sheet
    Rachel creates new sheet, names the sheet with just the smaple numb, then she copies and pastes the previous sheet into the newsheet
    so this includes the purity, all data. She puts in the sample number into B1.
    All of this is just her process, the master has no standard
    So then there's a 3rd excel file on the cloud(the network drive) so you paste the result of the equation into it
    this excel would be called let's say OVERALL_SAMPLE_DATA 
    it would have sample in column 1, purity in column 2, moisture in column 3, and other data
'''