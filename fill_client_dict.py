import pandas as pd

def main():
    perms = pd.read_excel('PERMS.xlsx')

    # print (perms)
    perm_dict = dict()
    for index, row in perms.iterrows():
        perm_dict.update({row['Customer Service']: {'Customer':row['Customer'],
                                                    'Origin':row['Origin'],
                                                    'Destinations':row['Destinations']}})
    
    return (perm_dict)