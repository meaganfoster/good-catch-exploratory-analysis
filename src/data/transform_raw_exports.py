import os
import pandas as pd
import shutil
import xlsxwriter
import openpyxl
import numpy as np
from pandas.api.types import is_string_dtype
from pandas.api.types import is_numeric_dtype
pd.set_option('display.width', 400)
pd.set_option('display.max_columns', 12)
# NOTE: you can choose to install/use external packages

# 0. import file
current_directory = os.getcwd()
# print(current_directory)
while True:
    file_name = input("Enter the name of the excel export. Do not include file extension (i.e. .xlsx).: ")
    goodcatch_filepath = (current_directory + "\\Raw Data Exports\\" + file_name + ".xlsx")
    print(goodcatch_filepath)

    if not(os.path.exists(goodcatch_filepath)):
        print("File does not exist. Enter new file name or press ESC to exit.")
    else:
        break
else:
    print("Reading file.")

# Add data to df data frame
df = pd.read_excel(goodcatch_filepath)
print("File successfully imported")
print("Please wait while we transform your file...")

# 1. Create GCID column to hold GC unique ID
df.insert(0, 'GCID', df.ID)
# print(df)


# 2. Update GCID to match the ID of the main GC record
# Create function to identify numeric values
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
    return False


# Update GCID to master ID where GCID is null
for i in range(len(df)):
    numeric_test = is_number(df.at[i, 'GCID'])
    if numeric_test and ~(df.isnull().at[i, 'GCID']):
        update_GCID = df.at[i, 'GCID']
    else:
        df.at[i, 'GCID'] = update_GCID


# 3. Create GC_Label to hold labels in ID column
df.insert(1, 'GC_Label', df.ID)

# Update GC_Label field to previous label
for i in range(len(df)):
    if ~(df.isnull().at[i, 'GC_Label']):
        update_GCID = df.at[i, 'GC_Label']
        # print(update_GCID)
    else:
        df.at[i, 'GC_Label'] = update_GCID

# Update Created At field to previous label
for i in range(len(df)):
    if ~(df.isnull().at[i, 'Created At']):
        update_GCID = df.at[i, 'Created At']
        # print(update_GCID)
    else:
        df.at[i, 'Created At'] = update_GCID

# 4. Drop blank rows
df_1 = df.drop(df[(df.ID == 'Causes') | (df.ID == 'Countermeasures') | (df.ID == 'Notes') | (df.ID == 'Reviewers') | (df.ID == 'Tags') | (df.ID == 'Teams') | (df.ID == 'What is Failed')].index)
final_df = df_1.dropna(subset=['GC_Label', 'FQID', 'QSC'])
print(final_df)


# 5. Create folder to store data

# Parent Directory path
parent_dir = os.getcwd() + "\Formatted Exports"
# print(parent_dir)
# mode
mode = 0o666


# Path; creates folder using file_name
path = os.path.join(parent_dir, file_name)
print("A new folder has been added, here: " + path)


# Create the directory
if os.path.exists(path):
    while True:
        remove_existing_files = input("Please confirm existing files have been renamed/removed. Type 'Y' to continue.:")
        if remove_existing_files not in ('Y'):
            print("Invalid value.")
        else:
            break
else:
    print("Creating directory...")
    os.mkdir(path)
    print("Directory '% s' created" % file_name)



# 6. Export dataframes to excel
def is_df_empty(df):
    return len(df.index) != 0

options = {}
options['strings_to_formulas'] = False
options['strings_to_urls'] = False

main_df = final_df[(final_df["GCID"] == final_df["ID"])]
writer_main = pd.ExcelWriter(current_directory + "\\Formatted Exports\\" + file_name + "\\" + file_name + "_main.xlsx", engine='xlsxwriter')
main_df.to_excel(writer_main, sheet_name='main', index=False)
writer_main.save()
# writer_main.close()

#print main exports only
# exit()

tags_df = final_df[(final_df["GC_Label"] == 'Tags')]
if is_df_empty(tags_df):
    tags_selected_columns = tags_df[["GCID", "GC_Label", "FQID", "Created At"]]
    tags_df_final = tags_selected_columns.copy()
    writer_tags = pd.ExcelWriter(current_directory + "\\Formatted Exports\\" + file_name + "\\" + file_name + "_tags.xlsx", engine='xlsxwriter')
    tags_df_final.to_excel(writer_tags, sheet_name='tags', index=False)
    writer_tags.save()
    # writer_tags.close()

what_is_failed_df = final_df[final_df["GC_Label"] == 'What is Failed']
if is_df_empty(what_is_failed_df):
    what_is_failed_selected_columns = what_is_failed_df[["GCID", "GC_Label", "FQID", "Created At"]]
    what_is_failed_df_final = what_is_failed_selected_columns.copy()
    writer_whatisfailed = pd.ExcelWriter(current_directory + "\\Formatted Exports\\" + file_name + "\\" + file_name + "_whatisfailed.xlsx", engine='xlsxwriter', options=options)
    what_is_failed_df_final.to_excel(writer_whatisfailed, sheet_name='whatisfailed', index=False)
    writer_whatisfailed.save()
    # writer_whatisfailed.close()

reviewers_df = final_df[final_df["GC_Label"] == 'Reviewers']
if is_df_empty(reviewers_df):
    reviewers_selected_columns = reviewers_df[["GCID", "GC_Label", "FQID", "Created At"]]
    reviewers_df_final = reviewers_selected_columns.copy()
    writer_reviewers = pd.ExcelWriter(current_directory + "\\Formatted Exports\\" + file_name + "\\" + file_name + "_reviewers.xlsx", engine='xlsxwriter', options=options)
    reviewers_df_final.to_excel(writer_reviewers, sheet_name='reviewers', index=False)
    writer_reviewers.save()
    # writer_reviewers.close()

causes_df = final_df[final_df["GC_Label"] == 'Causes']
if is_df_empty(causes_df):
    causes_selected_columns = causes_df[["GCID", "GC_Label", "FQID", "Created At"]]
    causes_df_final = causes_selected_columns.copy()
    writer_causes = pd.ExcelWriter(current_directory + "\\Formatted Exports\\" + file_name + "\\" + file_name + "_causes.xlsx", engine='xlsxwriter', options=options)
    causes_df_final.to_excel(writer_causes, sheet_name='causes', index=False)
    writer_causes.save()
    # writer_causes.close()

notes_df = final_df[final_df["GC_Label"] == 'Notes']
if is_df_empty(notes_df):
    notes_selected_columns = notes_df[["GCID", "GC_Label", "FQID", "Created At"]]
    notes_df_final = notes_selected_columns.copy()
    writer_notes = pd.ExcelWriter(current_directory + "\\Formatted Exports\\" + file_name + "\\" + file_name + "_notes.xlsx", engine='xlsxwriter', options=options)
    notes_df_final.to_excel(writer_notes, sheet_name='notes', index=False)
    writer_notes.save()
    # writer_notes.close()

teams_df = final_df[final_df["GC_Label"] == 'Teams']
if is_df_empty(teams_df):
    teams_selected_columns = teams_df[["GCID", "GC_Label", "FQID", "Created At"]]
    teams_df_final = teams_selected_columns.copy()
    writer_teams = pd.ExcelWriter(current_directory + "\\Formatted Exports\\" + file_name + "\\" + file_name + "_teams.xlsx", engine='xlsxwriter', options=options)
    teams_df_final.to_excel(writer_teams, sheet_name='teams', index=False)
    writer_teams.save()
    # writer_teams.close()

countermeasures_df = final_df[final_df["GC_Label"] == 'Countermeasures']
if is_df_empty(countermeasures_df):
    countermeasures_selected_columns = countermeasures_df[["GCID", "GC_Label", "FQID", "Created At"]]
    countermeasures_df_final = countermeasures_selected_columns.copy()
    writer_countermeasures = pd.ExcelWriter(current_directory + "\\Formatted Exports\\" + file_name + "\\" + file_name + "_countermeasures.xlsx", engine='xlsxwriter', options=options)
    countermeasures_df_final.to_excel(writer_countermeasures, sheet_name='countermeasures', index=False)
    writer_countermeasures.save()
    # writer_countermeasures.close()

print("Exports saved here: " + path)

# 6. Move original file to the new file directory
print("Please wait while we move your original file to this location.")

# shutil.move(current_directory + "\\Raw Data Exports\\" + file_name + ".xlsx", os.path.join(path))

print("Your original excel spreadsheet has been moved here: " + path + file_name)
print("Transformation complete. Your files are ready for use.")