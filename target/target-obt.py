############################################################################################
#
#   target-data processing rewrite for Orbitrap data
#   Last modified: 2026-07-10
#
############################################################################################
# Ensure you've installed all the dependencies from requirements.txt before proceeding.
# A rewrite of the original MATLAB script (Conrad Pritchard, v2.0.1, 3/4/2024) using pandas; rewritten by Shubhang Vyas.

#################
## --> Conventions
#################
# `IntS` *always* means "Internal Standard".
# `InjS` *always* means "Injection Standard".
# 
# `MB`: method blank
# `DB`: double blank
# `calstd`: calibration standard
# `aQC`: analytical QC samples -- CCV, ISC, LB, EPA, ICV, QC
# `mQC`: method QC samples -- LCS, LLLCS, LCSD, MB
# `PA`: peak area                                                   (compare to `PA` in the MATLAB script)
# `InjS_PA`: injection peak area                                    (compare to `injPA` in the MATLAB script)
# `IntS_PA`: internal standard peak area                            (compare to `ISPA` in the MATLAB script)
# `realconc`: real/actual concentration                             (compare to `actconc` in the MATLAB script)
# `compconc`: computed concentration                                (compare to `conc` in the MATLAB script)
# `real_IntS_conc`: real/actual internal standard concentration     (compare to `ISact` or `actISconc` in the MATLAB script)
# `sidx`: sample index                                              (compare to `rnum` in the MATLAB script)



#######################################################
### -- Import all the dependencies that we need. -- ###
#######################################################
print("Importing dependencies...")
try:
    import os
    import pandas as pd
    import numpy as np
    import openpyxl as pxlsx
    import re
# If we couldn't import a dependency, notify the user and exit with error code 1 (FAILURE).
except ImportError as e:
    print(f">>> Couldn't import module {e}! <<<")
    print("If the required modules aren't installed, consider running `pip install -r requirements.txt` to install them.")
    exit(1)


##########################################
### -- Prompt the user for options. -- ###
##########################################
##### 1) Get the target path.
# Get the file name from the user that we want to use.
file_name = input("Enter the name of the Excel or Text file (with extension) to process: ")
# If the user didn't give an extension, or gave a bad file extension, ask them to provide a valid one.
# re.search: if `file_name` does not end with `.xlsx`, `.xls`, or `.txt`, there was no valid extension.
if not re.search(r'\.(xlsx|xls|txt)$', file_name):
    ext = input("> Please provide a valid file extension (xlsx, xls, or txt): ")
    # If it's still not a valid extension, notify the user and exit with error code 1 (FAILURE).
    if ext not in ["xlsx", "xls", "txt"]:
        print(">>> Not a valid extension. <<<")
        exit(1)
    
    # Remove the invalid extension, if there was one.
    # re.sub: replace ".[bad_extension]" with "" (nothing), if it exists.
    file_name = re.sub(r'\.[^.]+$', '', file_name) + '.' + ext

# Split `file_name` into its base name and extension.
test_name, file_ext = os.path.splitext(file_name)

# Get the directory folder name that has this file.
dir_name = input("Enter the path of the directory with the file (or leave blank for current directory): ")
# If the user didn't provide a path, use the current working directory.
if not dir_name:
    print("> Using current working directory as the path.")
    dir_name = os.getcwd()

# Make the full path.
full_path = os.path.join(dir_name, file_name)

# If the path doesn't exist, notify the user and exit with error code 1 (FAILURE).
if not os.path.exists(full_path):
    print(f">>> Couldn't find a file at {full_path}! <<<")
    exit(1)

##### 2) Get the sample mass path, if specified.
nis_filename = input("Enter the name of the Excel file to set as the NIS file: ")
# If the user gave a bad file extension that is not .xlsx or .xls, notify the user and exit with error code 1 (FAILURE).
if re.search(r'\.(?!xlsx|xls$)', nis_filename):
    print(">>> Not a valid file extension for the NIS file -- must be .xlsx or .xls. <<<")
    exit(1)
    
# If there was no extension, add .xlsx to the end of the file name.
if not re.search(r'\.(xlsx|xls)$', nis_filename):
    nis_filename += ".xlsx"

# Get the directory folder name that has this file.
nis_dirname = input("Enter the path of the directory with the NIS file (or leave blank for current directory): ")
# If the user didn't provide a path, use the directory of the target file.
if not nis_dirname:
    print("> Using the same directory as the target file for the NIS file.")
    nis_dirname = dir_name

# Make the full path for the NIS file.
nis_path = os.path.join(nis_dirname, nis_filename)

# If the path doesn't exist, notify the user and exit with error code 1 (FAILURE).
if not os.path.exists(nis_path):
    print(f">>> Couldn't find a NIS file at {nis_path}! <<<")
    exit(1)

##### 2) Get options for including method blanks and injection type.
# Determine if we want to subtract method blanks from the samples.
# Keep asking the user until they give a valid response.
method = ""
while True:
    user_asked_method = input("Select a method (dust/soil/aqueous/plasma): ").strip().lower()
    if user_asked_method in ['dust', 'soil', 'aqueous', 'plasma']:
        method = user_asked_method
        break
    
######################################################
### -- Process the target data file into pandas -- ###
######################################################

if file_name.endswith(('.xlsx', '.xls')):
    # Read the Excel file into a pandas DataFrame.
    df = pd.read_excel(full_path, dtype=str)
elif file_name.endswith('.txt'):
    # Read the text file into a pandas DataFrame.
    df = pd.read_csv(full_path, sep='\t', dtype=str)

nis_df = pd.read_excel(nis_path, dtype=str)

# Make sure that the required columns that we need, exist.
target_columns_to_check = [
    'Filename',
    'Sample Type',
    'Compound',
    'Type',
    'RT',
    'Actual RT',
    'RT Delta',
    'Area',
    'Calculated Amt',
    'Theoretical Amt',
    'm/z (Expected)',
    'm/z (Apex)',
    'm/z (Delta)',
    'ISTD Response',
    'ISTD Amt',
    'Final Units'
]
NIS_columns_to_check = [
    'Internal Standard',
    'NIS Standard',
    'Compound'
]

for col in target_columns_to_check:
    # If the column is not in the DataFrame, notify the user and exit with error code 1 (FAILURE).
    if col not in df.columns:
        print(f">>> Missing required column '{col}' in the target data file! <<<")
        exit(1)
for col in NIS_columns_to_check:
    # If the column is not in the DataFrame, notify the user and exit with error code 1 (FAILURE).
    if col not in nis_df.columns:
        print(f">>> Missing required column '{col}' in the NIS file! <<<")
        exit(1)


#########################################################################
### -- Collect unique components and samples from the target data. -- ###
#########################################################################
# Collect unique components and samples.
first_sample = df.iloc[0]['Filename']
unique_samples = df['Filename'].unique().tolist()
unique_samples_rows = df[df['Filename'].isin(unique_samples)]
# Collect all the components from the first sample.
# The line below assumes that the first sample has all of the components that we want to analyze, and that they are all unique.
unique_components = df[df['Filename'] == first_sample]

# Collect our components of interest and internal standards.
components_of_interest = unique_components[unique_components['Type'] == 'Target Compound']
component_names = components_of_interest['Compound'].tolist()
# TODO: is this always true? Does this match the later check 'Type' == 'Internal Standard'?
all_IS_list = unique_components[unique_components['Type'] != 'Target Compound']

IS_list = pd.DataFrame(columns=all_IS_list.columns)
NIS_list = pd.DataFrame(columns=all_IS_list.columns)

for idx, comp in all_IS_list.iterrows():
    if comp['Compound'] in nis_df['Internal Standard'].values:
        IS_list = IS_list.append(comp)
    else:
        NIS_list = NIS_list.append(comp)

# Filters for extracting different kinds of samples.
is_calstd = unique_samples_rows['Sample Type'] == 'Cal Std'
# TODO: is_aQC was commented out in matlab script
# is_aQC = ~is_calstd & (unique_samples_rows['Sample Name'].str.contains('CCV|ISC|LB|ICV|AFFF|ALE'))
is_mQC = ~is_calstd & ~is_aQC & (unique_samples_rows['Sample Name'].str.contains('LSC|X|MB|QC'))
is_sample = ~is_calstd & ~is_aQC & ~is_mQC & ~(unique_samples_rows['Sample Name'].str.contains('DB'))

calstd = unique_samples_rows[is_calstd]
# TODO: aQC was commented out in matlab script
# aQC = unique_samples_rows[is_aQC]
mQC = unique_samples_rows[is_mQC]
# TODO: name conflict is confusing; rename?
samples = unique_samples_rows[is_sample]


################################################################
### -- Extract data and build pivots on the data we want. -- ###
################################################################
# TODO: can we guarantee that all "samples" all have components that are in components_of_interest?
# TODO: can we guarantee that a sample pivot-table, e.g. sample_ISPA, will have the same number of rows as components_of_interest?
def build_pivot(df, value_col, index_list=component_names, to_numeric=True):
    # First, obtain a unique (no repeat) order of our sample indices.
    col_order = df["Filename"].unique().tolist()

    # Convert our 2D dataframe into a pivot table.
    table = df.pivot_table(
        index="Compound",
        columns="Filename",
        values=value_column,    # Use values from the specific value column we want.
        aggfunc='first',        # If there are duplicates of our component name (which should never happen), use the first occurrence of that component.
        sort=False
    )
    
    # Reindex our sample indices based on the order of our index_list (by default, our component names).
    table = table.reindex(index=index_list, columns=col_order)

    # If we need to convert string values to numeric ones, do it.
    if to_numeric:
        table = table.apply(pd.to_numeric, errors='coerce')
    
    return table

# Build pivots for each data segment that we want.

# TODO: rough rewrite, check validity
if len(all_IS_list) > 0:
    index_list = all_IS_list['Compound'].tolist()
    # TODO: this excludes mQC/aQC from calculations; are the pivot tables the same?
    # TODO: does the mismatched index_list stretch the pivot table?
    IS_calstd = calstd[calstd['Type'] == 'Internal Standard']
    IS_sample = samples[samples['Type'] == 'Internal Standard']
    
    IS_calstd_realconc = build_pivot(IS_calstd, 'ISTD Amt', index_list=index_list)
    IS_calstd_PA = build_pivot(IS_calstd, 'Area', index_list=index_list)
    IS_sample_realconc = build_pivot(IS_sample, 'ISTD Amt', index_list=index_list)
    IS_sample_PA = build_pivot(IS_sample, 'Area', index_list=index_list)

intermediate_calstd = calstd[calstd['Type'] != 'Internal Standard']
# intermediate_aQC = aQC
intermediate_mQC = mQC
intermediate_samples = samples[samples['Type'] != 'Internal Standard']

calstd_name = build_pivot(intermediate_calstd, 'Compound', to_numeric=False)
calstd_realconc = build_pivot(intermediate_calstd, 'Theoretical Amt')
calstd_IS_realconc = build_pivot(intermediate_calstd, 'ISTD Amt')
calstd_ISPA = build_pivot(intermediate_calstd, 'ISTD Response')
calstd_compconc = build_pivot(intermediate_calstd, 'Calculated Amt')
calstd_PA = build_pivot(intermediate_calstd, 'Area')

# aQC_name = build_pivot(intermediate_aQC, 'Compound', to_numeric=False)
# aQC_realconc = build_pivot(intermediate_aQC, 'Theoretical Amt')
# aQC_IS_realconc = build_pivot(intermediate_aQC, 'ISTD Amt')
# aQC_ISPA = build_pivot(intermediate_aQC, 'ISTD Response')
# aQC_compconc = build_pivot(intermediate_aQC, 'Calculated Amt')
# aQC_PA = build_pivot(intermediate_aQC, 'Area')

mQC_name = build_pivot(intermediate_mQC, 'Compound', to_numeric=False)
mQC_realconc = build_pivot(intermediate_mQC, 'Theoretical Amt')
mQC_IS_realconc = build_pivot(intermediate_mQC, 'ISTD Amt')
mQC_ISPA = build_pivot(intermediate_mQC, 'ISTD Response')
mQC_compconc = build_pivot(intermediate_mQC, 'Calculated Amt')
mQC_PA = build_pivot(intermediate_mQC, 'Area')

sample_name = build_pivot(intermediate_samples, 'Compound', to_numeric=False)
# TODO: sample_realconc was not in the original matlab script
sample_IS_realconc = build_pivot(intermediate_samples, 'ISTD Amt')
sample_ISPA = build_pivot(intermediate_samples, 'ISTD Response')
sample_compconc = build_pivot(intermediate_samples, 'Calculated Amt')
sample_PA = build_pivot(intermediate_samples, 'Area')


##########################################
### -- Determine calibration range. -- ###
##########################################
sort_order = calstd_realconc.iloc[0, :].argsort()

# TODO: isn't calstd_realconc_sorted = sort_order, or am I stupid?
calstd_realconc_sorted = calstd_realconc.iloc[:, sort_order]
calstd_compconc_sorted = calstd_compconc.iloc[:, sort_order]
e = calstd_compconc_sorted / calstd_realconc_sorted

# `ee` is a bool matrix as to whether the element @ (i, j) in `e` is in the calibration range (0.7, 1.3).
lowlim = pd.Series(np.nan, index=range(e.shape[0]), dtype=float)
highlim = pd.Series(np.nan, index=range(e.shape[0]), dtype=float)
# TODO: is this right?
for idx, c in e.iterrows():
    # Iterate through each value in the row until we find the first value that is either NaN or outside the range (0.7, 1.3).
    for j, val in c.items():
        if pd.isna(val) or val < 0.7 or val > 1.3:
            lowlim[idx] = val
            break
    # Iterate through each value in the row in reverse order until we find the first value that is either NaN or outside the range (0.7, 1.3).
    for j, val in reversed(c.items()):
        if pd.isna(val) or val < 0.7 or val > 1.3:
            highlim[idx] = val
            break


#####################################################
### -- Determine IS and NIS Peak Area Recovery -- ###
#####################################################
if method in ['plasma', 'soil', 'dust']:
    sample_ISrec = pd.DataFrame(columns=sample_ISPA.columns, index=sample_ISPA.index)

    for idx, c in components_of_interest.iterrows():
        # find the NIS index for this component
        nis_idx = nis_df[nis_df['Compound'] == c['Compound']].index
        # If we couldn't find a NIS for this component, notify the user and exit with error code 1 (FAILURE).
        if len(nis_idx) == 0:
            print(f">>> Couldn't find a NIS for component {c['Compound']}! <<<")
            exit(1)

        nis_name = nis_df.loc[nis_idx[0], 'NIS Standard']
        IS_idx = all_IS_list[all_IS_list['Compound'] == nis_name].index
        if len(IS_idx) == 0:
            print(f">>> Couldn't find an IS for NIS {nis_name}! <<<")
            exit(1)
        
        nis_idx = nis_idx[0]
        IS_idx = IS_idx[0]

        # TODO: check validity
        sample_ISrec.loc[idx, :] = \
            sample_ISPA.loc[idx, :] \
            / IS_sample_PA.loc[IS_idx, :] \
            / (calstd_ISPA.loc[idx, :] / IS_calstd_PA.loc[IS_idx, :]).mean(axis=1, skipna=True) \
            * calstd_IS_realconc.loc[idx, :].mean(axis=1) / sample_IS_realconc.loc[idx, :]

    NIS_rec = pd.DataFrame(columns=IS_sample_PA.columns, index=IS_sample_PA.index)    
    for idx, c in NIS_list.iterrows():
        # TODO: wildly off from the matlab script + feels broken
        nis_name = nis_df[nis_df['Compound'] == c['Compound']].iloc[0]['NIS Standard']
        IS_idx = all_IS_list[all_IS_list['Compound'] == nis_name].index
        NIS_rec.iloc[idx, :] = IS_sample_PA.iloc[IS_idx, :] / IS_calstd_PA.iloc[IS_idx, :].mean(axis=1, skipna=True)

    for idx, c in IS_list.iterrows():
        IS_idx = nis_df[nis_df['Internal Standard'] == c['Compound']].index
        target_IS_idx = all_IS_list[all_IS_list['Compound'] == c['Compound']].index
        NIS_name = nis_df.iloc[IS_idx[0], 'NIS Standard']
        NIS_idx_in_IS = all_IS_list[all_IS_list['Compound'] == NIS_name].index
        IS_rec.iloc[idx, :] = \
            IS_sample_PA.iloc[target_IS_idx[0], :] \
            / IS_sample_PA.iloc[NIS_idx_in_IS[0], :] \
            / (IS_calstd_PA.iloc[target_IS_idx[0], :] / IS_calstd_PA.iloc[NIS_idx_in_IS[0], :]).mean(axis=1, skipna=True) \
            * IS_calstd_realconc.iloc[target_IS_idx[0], :].mean(axis=1) / IS_sample_realconc.iloc[target_IS_idx[0], :]

else:
    sample_ISrec = sample_ISPA / calstd_ISPA.mean(axis=1) * (calstd_IS_realconc.mean(axis=1) / sample_IS_realconc)
    ISrec = IS_sample_PA / IS_calstd_PA.mean(axis=1) * (IS_calstd_realconc.mean(axis=1) / IS_sample_realconc)


##################################################################
### -- Check method accuracy and determine reporting limits -- ###
##################################################################
# If there are method QC samples
if len(mQC) > 0:
    
    pass
else:
    pass

# TODO: LCMS analytical recovery is never used?