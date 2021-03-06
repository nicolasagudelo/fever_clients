import pandas as pd
from os import getcwd, startfile
from tkinter import filedialog, Tk,ttk
import fill_client_dict

# Normalized list will have the data we need from the csvs organized as we want it to be able to manipulate it.
normalized_list = []

#fill client dict will get a dictionary with all the Permanents with their clients, origins, and destinations.
perm_dict = fill_client_dict.main()

# Loop through the files corresponding to the last 4 weeks.
for i in range(0,4):
    # Create a window to ask the user where the files are located.
    root = Tk()

    frm = ttk.Frame(root, padding=10)
    frm.grid()
    ttk.Label(frm, text="Select file #{n}".format(n = i + 1)).grid(column=0, row=0)
    dirname = filedialog.askopenfilename(parent=root, initialdir=getcwd(),
                                        title='Please select file #{n}'.format(n = i + 1))

    # If the user closes the file dialog window we exit the script.
    if (len(dirname) == 0):
        print('Leaving program.')
        exit()

    # We close the window once we have what we need.
    root.destroy()

    ####################### USING PANDAS #####################

    # We set our dataframe to be equal to the csv the user selected.
    df = pd.read_csv(dirname)

    # At the moment the first row is always row number 6
    first_row = 6
    # We get the last row by looking for the index of the first row that has the text 'Generated on:' on the first column
    last_row = df[df.iloc[:, 0] == 'Generated on:'].index[0]

    pd.options.display.max_rows = 100
    pd.options.display.max_columns = 10

    # We get the info we want from the csv and pass it into a list, then we put that list into normalized list as a new item
    for j in range(first_row,last_row):
        tmp_lst = [df.iloc[j, 1], df.iloc[j, 3], df.iloc[j, 6], df.iloc[j, 5], str(i+1)]
        normalized_list.append(tmp_lst)

# Headers of our dataframe
header = ['Resolution', 'Permanent', 'Issue', 'Jira', 'Week']

df2 = pd.DataFrame(normalized_list).fillna(value="---> NO INFO <---")
df2.columns = header

# We print the dataframe with the info that is going to be used to the User so he can check one last time before creating the file.
print(df2)

input('\nThis is the information that will be used to generate the report.\nPress Enter if you want to continue')

# Create a list with only one instance of each permanent
permanents = []

for line in normalized_list:
    # We check if the permanent is already on the permanent list, if so we continue if not we add it
    if line[1] in permanents:
        continue
    # If not we add it to the list.
    permanents.append(line[1])

issue_list = []

for permanent in permanents:
    # Initialize our counters for each issue that we want to keep track of.
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
    for line in normalized_list:
        if line[0] == 'Duplicate':
            continue
        # We extract what was the issue for each line and in which week happened.
        issue = line[2]
        week = line[4]
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
df = pd.DataFrame(issue_list)

# Adding a header to our dataframe (df).
df.columns = header

# Creating our excel file.
while True:
    try:
        df.to_excel('output.xlsx',header=True, index= False)
    except PermissionError:
        input('Close the file before attempting to create a new one.\nPress Enter once you have closed the file to attempt to create it again.')
    
    break
        
# We tell the user where the file is located.
print("File generated on", getcwd())

input("Press Enter to open the file and close this window.")

# We try to open the file so the user can see it right away.
try:
    file = getcwd() + "\\output.xlsx"
    print ("Attempting to open file")
    startfile(file)
except FileNotFoundError:
    print("Unable to open file, try opening it normally.")
    exit()