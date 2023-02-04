import pandas as pd
import tkinter as tk
from tkinter import filedialog

def select_files():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="NOTE: The file must contain 2 sheets, a hierarchy pull and an IC schedule query" ,filetypes=[("Please select a file", "*.xlsx")])

    return file_path

def get_data(file_path):
    # read the excel file into a dataframe
    hierarchy_df = pd.read_excel(file_path, sheet_name=0)

    if hierarchy_df.columns[0] == "Hierarchy":
        schedule_df = pd.read_excel(file_path, sheet_name=1)
        hierarchy_df = pd.read_excel(file_path, sheet_name=0, usecols=['ENTITIES_Symbol','ENTITIES_Parent'])
    else:
        schedule_df = pd.read_excel(file_path, sheet_name=0)
        hierarchy_df = pd.read_excel(file_path, sheet_name=1, usecols=['ENTITIES_Symbol','ENTITIES_Parent'])
    
    # return the to dataframes
    return hierarchy_df, schedule_df

def find_leastCommonParent(hierarchy_df, sched_df):
    pair_ents = sched_df[['ENTITIES - Symbol Name', 'Offset - Symbol Name']]
    pair_ents = pair_ents.reset_index()

    root_ent = hierarchy_df['ENTITIES_Symbol'][0]

    leastCommonParents = []
    elimList = []

    for index, row in pair_ents.iterrows():
        #print(row['ENTITIES - Symbol Name'], row['Offset - Symbol Name'])
        offEnt_parents = []
        offEnt = row['Offset - Symbol Name']
        baseEnt = row['ENTITIES - Symbol Name']

        if offEnt == "TOTAL":
            leastCommonParents.append("")
            elimList.append("")
        else:
            while offEnt != root_ent:
                offEnt_parents.append(hierarchy_df[hierarchy_df['ENTITIES_Symbol']==offEnt].iloc[0,1])
                offEnt = hierarchy_df[hierarchy_df['ENTITIES_Symbol']==offEnt].iloc[0,1]

            while baseEnt not in offEnt_parents:
                baseEnt = hierarchy_df[hierarchy_df['ENTITIES_Symbol']==baseEnt].iloc[0,1]

            leastCommonParents.append(baseEnt)
            elimList.append(baseEnt + "E")

    sched_df['ENTITIES - Least Common Parent'] = leastCommonParents
    sched_df['ENTITIES - Elim Ents'] = elimList

    return sched_df


if __name__ == "__main__":

    file_path = select_files()

    hrchy_df, sch_df = get_data(file_path)

    updatedSched_df = find_leastCommonParent(hrchy_df, sch_df)

    saveFile_path = file_path.replace('.xlsx', '_Output.xlsx')

    updatedSched_df.to_excel(saveFile_path, sheet_name='Updated Schedule Query', index=False)
    
    print("DONE...")


