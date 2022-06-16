# Author: Nicolas Agudelo

import linecache
import pandas
import os
from tkinter import filedialog, Tk,ttk
import fill_client_dict

fixed_lines = []

perm_dict = fill_client_dict.main()

# Open the files

for i in range(0,4):

    root = Tk()

    frm = ttk.Frame(root, padding=10)
    frm.grid()
    ttk.Label(frm, text="Select file #{n}".format(n = i+1)).grid(column=0, row=0)
    dirname = filedialog.askopenfilename(parent=root, initialdir=os.getcwd(),
                                        title='Please select file #{n}'.format(n = i + 1))

    if (len(dirname) == 0):
        print('Leaving program.')
        exit()

    # Parse information of the file into lists

    root.destroy()

    lines = []

    # Define the range by taking the first line you want to take (inclusive) and the line after the last
    # you want to use (exclusive)

    first_line = 8 #int(input ("Write the start line to read the file\n"))
    last_line = int(input ("Write the last line to read the file\n"))
    last_line = last_line + 1

    for line in range(first_line, last_line, 1):
        particular_line = linecache.getline(dirname,line)
        particular_line = particular_line.replace(';',',')
        particular_line = particular_line.strip(',')
        particular_line = particular_line.strip('\n')
        particular_line = particular_line.strip(',,')
        particular_line = particular_line.replace(',,',',')
        if "Duplicate" in particular_line: continue
        #After formatting the line we added to the lines list
        lines.append(particular_line)

    # Splitting the strings in lines into lists

    for line in lines:
        list_line = line.split(",")
        list_line.pop()
        list_line.append(str(i+1))
        fixed_lines.append(list_line)

# Create a list with only one instance of each permanent on the file
permanents = []

for line in fixed_lines:
    # We check if the permanent is already on the permanent list, if so we continue if not we add it
    if line[1] in permanents:
        continue
    permanents.append(line[1])

# issue_list will be the list were we will have the info we ultimately wanted:
# Permanent name, equipment issues, link degradation, link outage.
issue_list = []

for permanent in permanents:
    equipment1 = 0
    link_degradation1 = 0
    link_outage1 = 0
    equipment2 = 0
    link_degradation2 = 0
    link_outage2 = 0
    equipment3 = 0
    link_degradation3 = 0
    link_outage3 = 0
    equipment4 = 0
    link_degradation4 = 0
    link_outage4 = 0
    total_equipment = 0
    total_degradation = 0
    total_outage = 0
    total = 0
    for line in fixed_lines:
        # We extract what was the issue for each line
        issue = line[3]
        week = line[8]
        # We check that the permanent we are checking is the one being affected in this line
        if permanent == line[1] and week == '1':# if the permanent we are checking matches the permanent in this line and corresponds to week 1
            match issue:
                # We check the issue if it's any of the ones we care about we add into the particular counter.
                case 'Equipment Problem': equipment1 += 1; total += 1; total_equipment +=1
                case 'Link Degradation': link_degradation1 += 1; total += 1; total_degradation +=1
                case 'Link Outage': link_outage1 += 1; total += 1; total_outage += 1
                case _: continue
        elif permanent == line[1] and week == '2':# if the permanent we are checking matches the permanent in this line and corresponds to week 2
            match issue:
                # We check the issue if it's any of the ones we care about we add into the particular counter.
                case 'Equipment Problem': equipment2 += 1; total += 1; total_equipment +=1
                case 'Link Degradation': link_degradation2 += 1; total += 1; total_degradation +=1
                case 'Link Outage': link_outage2 += 1; total += 1; total_outage += 1
                case _: continue
        elif permanent == line[1] and week == '3':# if the permanent we are checking matches the permanent in this line and corresponds to week 3
            match issue:
                # We check the issue if it's any of the ones we care about we add into the particular counter.
                case 'Equipment Problem': equipment3 += 1; total += 1; total_equipment +=1
                case 'Link Degradation': link_degradation3 += 1; total += 1; total_degradation +=1
                case 'Link Outage': link_outage3 += 1; total += 1; total_outage += 1
                case _: continue
        elif permanent == line[1] and week == '4':# if the permanent we are checking matches the permanent in this line and corresponds to week 3
            match issue:
                # We check the issue if it's any of the ones we care about we add into the particular counter.
                case 'Equipment Problem': equipment4 += 1; total += 1; total_equipment +=1
                case 'Link Degradation': link_degradation4 += 1; total += 1; total_degradation +=1
                case 'Link Outage': link_outage4 += 1; total += 1; total_outage += 1
                case _: continue
        else:
            continue
    try:
        customer = perm_dict[permanent]['Customer']
    except KeyError:
        for key, value in perm_dict.items():
            if key.startswith(permanent):
                customer = value['Customer']
                break
            else: customer = 'Customer not found, add manually'
    try:
        origin = perm_dict[permanent]['Origin']
    except KeyError:
        for key, value in perm_dict.items():
            if key.startswith(permanent):
                origin = value['Origin']
                break
            else: origin = 'Origin not found, add manually'
    try:
        destination = perm_dict[permanent]['Destinations'].rstrip(' // ').replace(' //',',\n')
    except KeyError:
        for key, value in perm_dict.items():
            if key.startswith(permanent):
                destination = value['Destinations'].rstrip(' // ').replace(' //',',\n')
                break
            else: destination = 'Destination not found, add manually'
    # We pass the information to our issue list with the final data.
    tmp_lst = [permanent, customer, origin, destination, equipment1, link_degradation1, link_outage1, equipment2, link_degradation2, link_outage2, equipment3, link_degradation3, link_outage3, equipment4, link_degradation4, link_outage4, total_equipment, total_degradation, total_outage, total]
    issue_list.append(tmp_lst)



# Generate dataframe from list and write to xlsx.

# Creating a header list to be the headers on the final excel file.
header = ['Affected Customer Circuit','Client', 'Origin', 'Destination', 'Equipment Problem','Link Degradation','Link Outage','Equipment Problem','Link Degradation','Link Outage','Equipment Problem','Link Degradation','Link Outage', 'Equipment Problem','Link Degradation','Link Outage','Equipment Problem','Link Degradation','Link Outage','Total']
# Passing the issue_list as the dataframe to be written into the excel file.
df = pandas.DataFrame(issue_list)

# Adding a header to our dataframe (df).
df.columns = header

# Creating our excel file.
while True:
    try:
        df.to_excel('output.xlsx',header=True, index= False)
    except PermissionError:
        input('Close the file before attempting to create a new one.\nPress Enter once you have closed the file to attempt to create it again.')
    
    break
        

print("File generated on", os.getcwd())




    



