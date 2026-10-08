#####
# Suspect Screening - Final Data Processing for LC-QTOF-MS
# Rewrite of v4.3 in MATLAB, Conrad Pritchard 3/19/2025

# STANDALONE SINGLE-PASS VERSION
# unfinished!
# TODO: add in the options for reading excel files that the matlab script does

import os
import sys
import re
import pandas as pd
import openpyxl
import math
import numpy as np

suspect_filename = input("Enter the name of the suspect file (with extension): ")
suspect_dirname = input("Enter the name of the directory where the file is located: ")
suspect_pathname = os.path.join(suspect_dirname, suspect_filename)

target_filename = input("Enter the name of the target file (with extension): ")
target_dirname = input("Enter the name of the directory where the file is located: ")
target_pathname = os.path.join(target_dirname, target_filename)

calibrant_filename = input("Enter the name of the calibrant file (with extension): ")
calibrant_dirname = input("Enter the name of the directory where the file is located: ")
calibrant_pathname = os.path.join(calibrant_dirname, calibrant_filename)

# Define constants.
c_me = 10  # Mass error threshold (ppm)
c_w50 = 1  # Width at 50% threshold
c_area = 1e3  # Area threshold
# TODO: not used in the latest version, v4.3.2; remove?
c_sn = 10  # Signal-to-noise threshold
c_h = 50  # Height threshold
c_q = 0.3 # Quality threshold
c_b = 0.1 # Baseline delta / height threshold

# Read the files from Excel into pandas DataFrames, and check that the required columns exist.
suspect_df = pd.read_excel(suspect_pathname)
target_df = pd.read_excel(target_pathname)

sample_req_cols = ['Sample Name', 'Sample Index', 'Injection Volume', 'Component Name', 'Component Index', 'Component Group Name', 'Area', 'Height', 'Quality', 'Retention Time', 'Width at 50%', 'Signal / Noise', 'Baseline Delta / Height', 'Formula', 'Precursor Mass', 'Found At Mass', 'Mass Error (ppm)', 'Library Hit', 'Library Score', 'Combined Score', 'Points Across Half Height']
target_req_cols = ['Sample Name', 'Sample Index', 'Injection Volume', 'Component Name', 'Component Group Name', 'Component Type', 'Retention Time', 'Retention Time Delta (min)', 'Precursor Mass', 'Mass Error Confidence', 'Mass Error (ppm)', 'IS Name', 'Area', 'IS Area', 'Actual Concentration', 'IS Actual Concentration', 'Calculated Concentration', 'Sample Type', 'Used']

for col in sample_req_cols:
    if col not in suspect_df.columns:
        print(f"Error: Required column '{col}' not found in suspect file.")
        sys.exit(1)
for col in target_req_cols:
    if col not in target_df.columns:
        print(f"Error: Required column '{col}' not found in target file.")
        sys.exit(1)

##############################################
## --- Determine the method of sampling --- ##
##############################################
# Get the first row's injection volume; if it's 100, our sampling method is SOIL; otherwise, AQUEOUS.
inj_vol = target_df.iloc[0]['Injection Volume']
method = 'SOIL' if inj_vol == 100 else 'AQUEOUS'

### -- [intern] #4 on matlab script; determine unique by component name
s_unique_comp = suspect_df['Component Name'].unique().tolist()
# TODO: practically unused
s_num_unique_comp = len(s_unique_comp)
# get first 1:s_num_unique_comp rows of suspect_df and collect their rows
# TODO: isn't this the same as getting all rows of the first sample name? then suspect_df[suspect_df['Sample Name'] == suspect_df.iloc[0]['Sample Name']]
s_compdf = suspect_df.head(s_num_unique_comp)

s_unique_samples = suspect_df['Sample Name'].unique().tolist()
s_num_unique_samples = len(s_unique_samples)

# create s_comp_grouplist: convert all rows of s_compdf into a dataframe where every row is [component name, component group name, formula, precursor mass]
s_comp_grouplist = s_compdf[['Component Name', 'Component Group Name', 'Formula', 'Precursor Mass']].values.tolist()

t_first_sample_filter = target_df['Sample Name'] == target_df.iloc[0]['Sample Name']
t_of_first_sample = target_df[t_first_sample_filter]
t_first_sample_targets = t_of_first_sample[t_of_first_sample['Component Type'].isin(['Quantifiers', 'Qualifiers'])]
t_component_name_list = t_first_sample_targets[['Component Name', 'IS Name', 'Precursor Mass']].values.tolist()
t_num_IS = len(t_first_sample_targets[t_first_sample_targets['Component Type'].isin(['Internal Standard', 'Internal Standards'])])

### -- [intern] #5 on matlab script
# for each component, extract specific data; if they match, append it to s_compfinal.
s_compfinal = pd.DataFrame(columns=['Component Data', 'Index', 'Calibrant Name', 'Component Group Name'])
for idx, comp in s_compdf.iterrows():
    # Get all rows in sample_df that match the current component name.
    matching_rows = suspect_df[suspect_df['Component Name'] == comp['Component Name']]
    # Iterate through all, and extract data for each row. If any of the rows are of interest, append to s_compfinal and break.
    for _, row in matching_rows.iterrows():
        # Extract relevant data from the row.
        mass_error = row['Mass Error (ppm)']
        width_50 = row['Width at 50%']
        area = row['Area']
        signoise = row['Signal / Noise']
        height = row['Height']
        quality = row['Quality']
        bdh = row['Baseline Delta / Height']

        # check if reqs met and sample name doesn't contain 'AFFF'
        if math.fabs(mass_error) < c_me and width_50 < c_w50 and area > c_area and signoise > c_sn and height > c_h and quality > c_q and bdh < c_b \
            and row['Sample Name'].find('AFFF') == -1:
            s_compfinal = s_compfinal.append([
                comp,
                idx,
                None,
                s_comp_grouplist[idx][1],  # Component Group Name
            ], ignore_index=True)
            break

### -- [intern] #6 on matlab script
# make empty vars; 2D matrix of [final component index, sample num]
s_compfinal_RT_all = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(s_num_unique_samples))
s_compfinal_area_all = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(s_num_unique_samples))
s_compfinal_masserr = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(s_num_unique_samples))
s_compfinal_libscore = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(s_num_unique_samples))
s_numhomolog = pd.Series(0, index=range(len(s_compfinal)))

for idx, comp in s_compfinal.iterrows():
    name = comp[0]['Component Name']
    all_rows = suspect_df[suspect_df['Component Name'] == name]
    
    for _, row in all_rows.iterrows():
        masserr = row['Mass Error (ppm)']
        width_50 = row['Width at 50%']
        area = row['Area']
        signoise = row['Signal / Noise']
        height = row['Height']
        quality = row['Quality']
        bdh = row['Baseline Delta / Height']
        
        if math.fabs(masserr) < c_me and width_50 < c_w50 and area > c_area and signoise > c_sn and height > c_h and quality > c_q and bdh < c_b:
            s_compfinal_RT_all.iloc[idx, row['Sample Name']] = row['Retention Time']
            s_compfinal_area_all.iloc[idx, row['Sample Name']] = area
            s_compfinal_masserr.iloc[idx, row['Sample Name']] = masserr
            s_compfinal_libscore.iloc[idx, row['Sample Name']] = row['Library Score']
        else:
            s_compfinal_RT_all.iloc[idx, row['Sample Name']] = 0
            s_compfinal_area_all.iloc[idx, row['Sample Name']] = 0
            s_compfinal_masserr.iloc[idx, row['Sample Name']] = 0
            s_compfinal_libscore.iloc[idx, row['Sample Name']] = 0
    
    # matlab: scomfinal_RT_allnz = find(scomfinal_RT_all(i,:)~=0);
    # get median of the non-zero RTs
    s_compfinal_RT_allnz = s_compfinal_RT_all.iloc[idx][s_compfinal_RT_all.iloc[idx] != 0].index.tolist()
    s_compfinal_RT = s_compfinal_RT_all.iloc[idx][s_compfinal_RT_allnz].median()
    
    # get ALL area values and find the max; we don't care about nonzero because we're getting the maximum of all values.
    s_compfinal_area = s_compfinal_area_all.iloc[idx].max()
    
    # get median of all non-zero mass error values
    s_compfinal_masserr_allnz = s_compfinal_masserr.iloc[idx][s_compfinal_masserr.iloc[idx] != 0].index.tolist()
    s_compfinal_masserr = s_compfinal_masserr.iloc[idx][s_compfinal_masserr_allnz].median()

    # get all lib score values and get max
    s_compfinal_libscore = s_compfinal_libscore.iloc[idx].max()
    
    # Find the number of compounds that have the same component group name as the current compound.
    s_numhomolog.iloc[idx] = len(s_compfinal[s_compfinal['Component Group Name'] == comp[3]])

### --- [intern] empty #6??
### --- [intern] #7 on matlab script
# get first 32 chars of suspect list file name; or all of it if the name is less than 32 chars.
trimmed_suspect_filename = suspect_filename[:32] if len(suspect_filename) > 32 else suspect_filename

# read calibrant match file
calmatch_df = pd.read_excel(calibrant_pathname, sheet_name=suspect_filename)

# for each final suspect component, if there's a calibrant match, then replace the third value from None to the calibrant name; otherwise, set it to 'NO_CALIBRANT'.
NO_CALIBRANT = 'NO_CALIBRANT'
for idx, comp in enumerate(s_compfinal):
    # check the first column to match the name for each row
    # TODO: is this right?
    match = calmatch_df[calmatch_df.iloc[:, 0] == comp[0]['Component Name']]
    
    if not match.empty:
        s_compfinal[idx][2] = match.iloc[0, 1]  # set the calibrant name
    else:
        s_compfinal[idx][2] = NO_CALIBRANT

### --- [intern] #8 on matlab script
# same thing as target filter script

# remember that the index list is based on the index of the component (this target row) in the target component name list
def build_pivot(df, value_col, idx_list=t_component_name_list, to_numeric=True):
    # Create a unique (no repeat) order of our sample indices.
    col_order = df["Sample Index"].unique().tolist()

    # Convert our 2D dataframe into a pivot table.
    table = df.pivot_table(
        index="Component Name",
        columns="Sample Index",
        values=value_col,
        aggfunc="first",
        sort=False
    )

    # Reindex our sample indices based on the order of our idx list.
    table = table.reindex(index=idx_list, columns=col_order)

    if to_numeric:
        table = table.apply(pd.to_numeric, errors='coerce')

    return table

# Filters for extracting rows out of the target data list.
# note that is_intstd and is_skip are not selected on and are instead used as filters for the other filters.
is_intstd = target_df["Component Type"].isin(["Internal Standard", "Internal Standards"])
is_calstd = ~is_intstd & target_df["Sample Type"] == "Standard"
is_aQC = ~is_intstd & ~is_calstd & re.search(r'CCV|ISC|LB|EPA|AFFF|QC', target_df["Sample Name"])
is_mQC = ~is_intstd & ~is_calstd & ~is_aQC & re.search(r'LCS|MB', target_df["Sample Name"])
is_skip = ~is_intstd & ~is_calstd & ~is_aQC & ~is_mQC & re.search(r'DB|blank check', target_df["Sample Name"])
is_sample = ~is_intstd & ~is_calstd & ~is_aQC & ~is_mQC & ~is_skip

target_calstd = target_df[is_calstd]
target_aQC = target_df[is_aQC]
target_mQC = target_df[is_mQC]
target_sample = target_df[is_sample]

target_calstd_IS_realconc = build_pivot(target_calstd, "IS Actual Concentration")
target_calstd_realconc = build_pivot(target_calstd, "Actual Concentration")
target_calstd_ISPA = build_pivot(target_calstd, "IS Area")
target_calstd_calcconc = build_pivot(target_calstd, "Calculated Concentration")
target_calstd_PA = build_pivot(target_calstd, "Area")
target_calstd_sidx = build_pivot(target_calstd, "Sample Index", to_numeric=False)
target_calstd_used = build_pivot(target_calstd, "Used", to_numeric=False)

target_aQC_name = build_pivot(target_aQC, "Sample Name", to_numeric=False)
target_aQC_IS_realconc = build_pivot(target_aQC, "IS Actual Concentration")
target_aQC_ISPA = build_pivot(target_aQC, "IS Area")
target_aQC_compconc = build_pivot(target_aQC, "Calculated Concentration")
target_aQC_PA = build_pivot(target_aQC, "Area")
target_aQC_sidx = build_pivot(target_aQC, "Sample Index", to_numeric=False)

target_mQC_name = build_pivot(target_mQC, "Sample Name", to_numeric=False)
target_mQC_IS_realconc = build_pivot(target_mQC, "IS Actual Concentration")
target_mQC_ISPA = build_pivot(target_mQC, "IS Area")
target_mQC_compconc = build_pivot(target_mQC, "Calculated Concentration")
target_mQC_PA = build_pivot(target_mQC, "Area")
target_mQC_sidx = build_pivot(target_mQC, "Sample Index", to_numeric=False)

target_sample_name = build_pivot(target_sample, "Sample Name", to_numeric=False)
target_sample_IS_realconc = build_pivot(target_sample, "IS Actual Concentration")
target_sample_ISPA = build_pivot(target_sample, "IS Area")
target_sample_compconc = build_pivot(target_sample, "Calculated Concentration")
target_sample_PA = build_pivot(target_sample, "Area")
target_sample_sidx = build_pivot(target_sample, "Sample Index", to_numeric=False)

### --- [intern] #9 on matlab script
# TODO: removed this before argsort: `.apply(pd.to_numeric, errors='coerce')`
sort_order = target_calstd_realconc.iloc[0, :].argsort()

# sort the lists we want
target_calstd_used_sorted = target_calstd_used.iloc[:, sort_order]

target_calstd_IS_realconc_sorted = target_calstd_IS_realconc.iloc[:, sort_order]
target_calstd_realconc_sorted = target_calstd_realconc.iloc[:, sort_order]
target_calstd_PA_sorted = target_calstd_PA.iloc[:, sort_order]
target_calstd_ISPA_sorted = target_calstd_ISPA.iloc[:, sort_order]

e = target_calstd_compconc_sorted / target_calstd_realconc_sorted

slopes = pd.Series(0, index=range(len(e)), dtype=float)
min_cals = pd.Series(0, index=range(len(e)), dtype=float)
max_cals = pd.Series(0, index=range(len(e)), dtype=float)

# for each target component in our calstd
for idx, comp in e.iterrows():
    # TODO: ???
    target_calcurve_x = pd.Series(dtype=float)
    target_calcurve_y = pd.Series(dtype=float)

    # For every calibration level, check if the conditions are met; if so, append into the x/y values from the actual concentration and peak area
    for col in e.columns:
        accuracy = e.loc[idx, col]
        if accuracy > 0.7 and accuracy < 1.3 and target_calstd_used_sorted.loc[idx, col].lower() == 'true':
            target_calcurve_x = target_calcurve_x.append(
                target_calstd_realconc_sorted.loc[idx, col] / target_calstd_IS_realconc_sorted.loc[idx, col],
                ignore_index=True
            )
            target_calcurve_y = target_calcurve_y.append(
                target_calstd_PA_sorted.loc[idx, col] / target_calstd_ISPA_sorted.loc[idx, col],
                ignore_index=True
            )
            
    if len(target_calcurve_x) > 2:
        # use a weight option of 1/x^2 for the linear regerssion
        # get the curve and the error, then extract the slope, min and max calibration levels
        # matlab: [calcurve, cal_error] = fit(tcalcurve_x(:), tcalcurve_y(:), 'poly1', fitOptions) where fitOptions = fitoptions('Weights', 1./(tcalcurve_x.^2))
        fit = np.polyfit(target_calcurve_x, target_calcurve_y, 1, w=1/target_calcurve_x**2)
        slope = fit[0]
        # TODO: this was adapted differently from the matlab; check validity
        min_cal = target_calcurve_x.min() * target_calstd_IS_realconc_sorted.loc[idx, :].min()
        max_cal = target_calcurve_x.max() * target_calstd_IS_realconc_sorted.loc[idx, :].max()

        slopes.loc[idx] = slope
        min_cals.loc[idx] = min_cal
        max_cals.loc[idx] = max_cal

### --- [intern] #10 on matlab script
s_IV_conc = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(s_compfinal_area_all.shape[1]))
s_response_factor = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_lowercal = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_uppercal = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
for idx, comp in s_compfinal.iterrows():
    calibrant_name = comp[2]
    target_calstd_idx = t_component_name_list[t_component_name_list['Component Name'] == calibrant_name].index.tolist()
    if calibrant_name == NO_CALIBRANT or not target_calstd_idx:
        print(f"Warning: No calibrant found for component '{comp[0]['Component Name']}'. Skipping response factor calculation.")
        s_IV_conc.iloc[idx, :] = np.nan
        s_response_factor.iloc[idx] = np.nan
        s_lowercal.iloc[idx] = np.nan
        s_uppercal.iloc[idx] = np.nan
    
    target_calstd_idx = target_calstd_idx[0]
    t_comp_mass = t_component_name_list.loc[target_calstd_idx, 'Precursor Mass']
    s_response_factor.iloc[idx] = slopes.loc[target_calstd_idx]
    s_lowercal.iloc[idx] = min_cals.loc[target_calstd_idx]
    s_uppercal.iloc[idx] = max_cals.loc[target_calstd_idx]

    # for each sample in our suspect data set
    for name in s_unique_samples:
        # Skip AFFF samples.
        if 'AFFF' in name:
            continue
        
        # If this sample is a sample in the target data set,
        if name in target_sample_name.columns:
            # remember: table is indexed by [component name, sample index];
            # component name: comp[0]['Component Name']
            # sample index: target_sample_sidx.loc[comp[0]['Component Name'], name]
            t_IS_realconc = target_sample_IS_realconc.loc[comp[0]['Component Name'], target_sample_sidx.loc[comp[0]['Component Name'], name]]
            t_ISPA = target_sample_ISPA.loc[comp[0]['Component Name'], target_sample_sidx.loc[comp[0]['Component Name'], name]]
        # If this sample is an aQC sample in the target data set,
        elif name in target_aQC_name.columns:
            t_IS_realconc = target_aQC_IS_realconc.loc[comp[0]['Component Name'], target_aQC_sidx.loc[comp[0]['Component Name'], name]]
            t_ISPA = target_aQC_ISPA.loc[comp[0]['Component Name'], target_aQC_sidx.loc[comp[0]['Component Name'], name]]
        # If this sample is a mQC sample in the target data set,
        elif name in target_mQC_name.columns:
            t_IS_realconc = target_mQC_IS_realconc.loc[comp[0]['Component Name'], target_mQC_sidx.loc[comp[0]['Component Name'], name]]
            t_ISPA = target_mQC_ISPA.loc[comp[0]['Component Name'], target_mQC_sidx.loc[comp[0]['Component Name'], name]]
            
        deref_idx = comp[1] # Get the component's index in s_comp_grouplist
        s_IV_conc.loc[idx, name] = t_IS_realconc * s_compfinal_area_all.loc[idx, name] \
            / (t_ISPA * slopes.loc[target_calstd_idx] * t_comp_mass / s_comp_grouplist.loc[deref_idx, 'Precursor Mass'])

### --- [intern] #11 on matlab script
# find all LB and MB and put into a QC
suspect_QC_df = suspect_df[suspect_df['Sample Name'].str.contains('LB|MB')]
s_low_RL = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_high_RL = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
if not suspect_QC_df.empty:
    # For each final suspect component:
    for idx, comp in s_compfinal.iterrows():
        target_calstd_idx = t_component_name_list[t_component_name_list['Component Name'] == comp[2]].index.tolist()
        if not target_calstd_idx:
            s_low_RL.iloc[idx] = np.nan
            s_high_RL.iloc[idx] = np.nan
        s_low_RL.iloc[idx] = max(3 * s_IV_conc.loc[idx, suspect_QC_df].max(), s_lowercal.loc[idx])
        s_high_RL.iloc[idx] = s_uppercal.loc[idx]
    
    # filter s_IV_conc, s_compfinal_area_all, and the sample names to remove any samples
    # that were in suspect_QC_df or were AFFF samples.
    cols_to_filter = suspect_QC_df['Sample Name'].tolist() + [name for name in s_unique_samples if 'AFFF' in name]
    s_IV_conc2 = s_IV_conc.drop(columns=cols_to_filter, errors='ignore')
    s_compfinal_area_all2 = s_compfinal_area_all.drop(columns=cols_to_filter, errors='ignore')
    s_unique_samples2 = [name for name in s_unique_samples if name not in cols_to_filter]
else:
    # According to the matlab script, this branch is NOT recommended.
    for idx, comp in s_compfinal.iterrows():
        s_low_RL.iloc[idx] = s_lowercal.loc[idx]
        s_high_RL.iloc[idx] = s_uppercal.loc[idx]
    
    cols_to_filter = [name for name in s_unique_samples if 'AFFF' in name]
    s_IV_conc2 = s_IV_conc.drop(columns=cols_to_filter, errors='ignore')
    s_compfinal_area_all2 = s_compfinal_area_all.drop(columns=cols_to_filter, errors='ignore')
    s_unique_samples2 = [name for name in s_unique_samples if name not in cols_to_filter]

### --- [intern] #12 on matlab script
if method == 'SOIL':
    # read in the soil masses
    soilmass_path = os.path.join(target_dirname, 'soil mass.xlsx')
    raw_soilmasses = pd.read_excel(soilmass_path, sheet_name='Sheet')
    # get the matrix mass data
    target_matrixmass = raw_soilmasses[:, 1] * (1 - raw_soilmasses[:, 2])
    # For each suspect sample, determine the matrix concentration and reporting limits
    s_matrix_sample_conc = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)))
    s_matrix_max_conc = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)))
    s_matrix_min_conc = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)))
    s_matrix_conc_LoQ = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)), dtype=str)
    s_matrix_conc_LoQ_undil = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)), dtype=str)
    for sidx in range(len(s_unique_samples2)):
        name = s_unique_samples2[sidx]
        mass_idx = raw_soilmasses[raw_soilmasses[0] == name].index
        if mass_idx.empty:
            print(f"Error: cannot find sample '{name}' in soil mass file.")
            sys.exit(1)
        
        # Determine matrix extraction factor.
        if raw_soilmasses.loc[sidx, 3] == 'soil':
            mat_exfactor = 0.4/0.1 * 1.5/1000
        elif raw_soilmasses.loc[sidx, 3] == 'dust':
            mat_exfactor = 0.6/0.45 * 1.5/1000 * 5/0.25
        elif raw_soilmasses.loc[sidx, 3] == 'SPE' or raw_soilmasses.loc[sidx, 3] == 'water':
            mat_exfactor = 0.75/0.375 * 5/1000*1000
        elif raw_soilmasses.loc[sidx, 3] == 'SPEsoil':
            mat_exfactor = 0.75/0.375 * 5/1000 * 32/2.5
        else:
            print(f"Error: unknown matrix type '{raw_soilmasses.loc[sidx, 3]}' for sample '{name}'.")
            sys.exit(1)
        
        common_factor = mat_exfactor / target_matrixmass[sidx]

        # Calculate matrix concentration
        s_matrix_sample_conc[:, sidx] = s_IV_conc2.iloc[:, sidx] * common_factor
        # Determine matrix upper and lower Loq range
        s_matrix_max_conc[:, sidx] = s_high_RL * common_factor
        s_matrix_min_conc[:, sidx] = s_low_RL * common_factor

        # [#13] correct diluted values; find instances of "_d##x" in the sample name; if so, extract the dilution factor
        match = re.search(r'_d(\d+)x', name)
        if match:
            dilution_factor = int(match.group(1))
        else:
            dilution_factor = 1
        # Replace values outside of the matrix LoQ
        for cidx in range(len(s_compfinal)):
            if s_matrix_sample_conc.iloc[cidx, sidx] > s_matrix_max_conc.iloc[cidx, sidx]:
                s_matrix_conc_LoQ.iloc[cidx, sidx] = f">{(s_matrix_max_conc.iloc[cidx, sidx] * dilution_factor):.2f}"
                s_matrix_conc_LoQ_undil.iloc[cidx, sidx] = f">{(s_matrix_max_conc.iloc[cidx, sidx]):.2f}"
            elif s_matrix_sample_conc.iloc[cidx, sidx] < s_matrix_min_conc.iloc[cidx, sidx]:
                s_matrix_conc_LoQ.iloc[cidx, sidx] = f"<{(s_matrix_min_conc.iloc[cidx, sidx] * dilution_factor):.2f}"
                s_matrix_conc_LoQ_undil.iloc[cidx, sidx] = f"<{(s_matrix_min_conc.iloc[cidx, sidx]):.2f}"
            else:
                s_matrix_conc_LoQ.iloc[cidx, sidx] = f"{(s_matrix_sample_conc.iloc[cidx, sidx] * dilution_factor):.2f}"
                s_matrix_conc_LoQ_undil.iloc[cidx, sidx] = f"{(s_matrix_sample_conc.iloc[cidx, sidx]):.2f}"

    units = 'ng/g'
else:
    # Calculate the matrix concentration
    # TODO: this was 1.5 / 0.87 in a completely different matlab/python script that did something similar to this; check validity.
    water_factor = 1.5 / 0.9
    s_matrix_sample_conc = s_IV_conc2 * water_factor
    s_matrix_min_conc = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)))
    s_matrix_max_conc = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)))
    s_matrix_conc_LoQ = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)), dtype=str)
    s_matrix_conc_LoQ_undil = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)), dtype=str)
    
    for sidx in range(len(s_unique_samples2)):
        s_matrix_max_conc.iloc[:, sidx] = s_high_RL * water_factor
        s_matrix_min_conc.iloc[:, sidx] = s_low_RL * water_factor
        # Extract dilution factor from sample name
        match = re.search(r'_d(\d+)x', s_unique_samples2.iloc[sidx])
        if match:
            dilution_factor = int(match.group(1))
        else:
            dilution_factor = 1

        for cidx in range(len(s_compfinal)):
            if s_matrix_sample_conc.iloc[cidx, sidx] > s_matrix_max_conc.iloc[cidx, sidx]:
                s_matrix_conc_LoQ.iloc[cidx, sidx] = f">{(s_matrix_max_conc.iloc[cidx, sidx] * dilution_factor):.2f}"
                s_matrix_conc_LoQ_undil.iloc[cidx, sidx] = f">{(s_matrix_max_conc.iloc[cidx, sidx]):.2f}"
            elif s_matrix_sample_conc.iloc[cidx, sidx] < s_matrix_min_conc.iloc[cidx, sidx]:
                s_matrix_conc_LoQ.iloc[cidx, sidx] = f"<{(s_matrix_min_conc.iloc[cidx, sidx] * dilution_factor):.2f}"
                s_matrix_conc_LoQ_undil.iloc[cidx, sidx] = f"<{(s_matrix_min_conc.iloc[cidx, sidx]):.2f}"
            else:
                s_matrix_conc_LoQ.iloc[cidx, sidx] = f"{(s_matrix_sample_conc.iloc[cidx, sidx] * dilution_factor):.2f}"
                s_matrix_conc_LoQ_undil.iloc[cidx, sidx] = f"{(s_matrix_sample_conc.iloc[cidx, sidx]):.2f}"

    units = 'ng/L'

### --- [intern] #13 on matlab script
# completed above

### --- [intern] #14 on matlab script
# there was no step 14?

### --- [intern] #15 on matlab script
# separate target compounds
s_matrix_conc_LoQ_tar = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)), dtype=str)
s_matrix_sample_conc_tar = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)))
s_compfinal_tar = pd.Series("", index=range(len(s_compfinal)), dtype=str)
s_compfinal_calstd_tar = pd.Series("", index=range(len(s_compfinal)), dtype=str)
s_compfinal_mass_tar  = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_compfinal_group_tar = pd.Series("", index=range(len(s_compfinal)), dtype=str)
s_compfinal_formula_tar = pd.Series("", index=range(len(s_compfinal)), dtype=str)
s_compfinal_RT_tar = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_masserr_tar = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_numhomolog_tar = pd.Series(0, index=range(len(s_compfinal)), dtype=int)
s_libscore_tar = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_IVconc_tar = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)))
s_compfinal_area_tar = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)))
s_response_factor_tar = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_lowercal_tar = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_uppercal_tar = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_min_matrixconc_tar = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_max_matrixconc_tar = pd.Series(0, index=range(len(s_compfinal)), dtype=float)

s_matrix_conc_LoQ_final = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)), dtype=str)
s_matrix_sample_conc_final = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)))
s_compfinal_final = pd.Series("", index=range(len(s_compfinal)), dtype=str)
s_compfinal_calstd_final = pd.Series("", index=range(len(s_compfinal)), dtype=str)
s_compfinal_mass_final = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_compfinal_group_final = pd.Series("", index=range(len(s_compfinal)), dtype=str)
s_compfinal_formula_final = pd.Series("", index=range(len(s_compfinal)), dtype=str)
s_compfinal_RT_final = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_masserr_final = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_numhomolog_final = pd.Series(0, index=range(len(s_compfinal)), dtype=int)
s_libscore_final = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_IVconc_final = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)))
s_compfinal_area_final = pd.DataFrame(0, index=range(len(s_compfinal)), columns=range(len(s_unique_samples2)))
s_response_factor_final = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_lowercal_final = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_uppercal_final = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_min_matrixconc_final = pd.Series(0, index=range(len(s_compfinal)), dtype=float)
s_max_matrixconc_final = pd.Series(0, index=range(len(s_compfinal)), dtype=float)

for idx, comp in s_compfinal.iterrows():
    # if the suspect component is alsoa target component
    if comp[0]['Component Name'] in t_component_name_list['Component Name'].values:
        s_matrix_conc_LoQ_tar.iloc[idx, :] = s_matrix_conc_LoQ.iloc[idx, :]
        s_matrix_sample_conc_tar.iloc[idx, :] = s_matrix_sample_conc.iloc[idx, :]
        s_compfinal_tar.iloc[idx] = comp[0]['Component Name']
        s_compfinal_calstd_tar.iloc[idx] = comp[2]
        s_compfinal_mass_tar.iloc[idx] = s_comp_grouplist[idx][3]
        # TODO: this is inefficient, do comp[3] or comp[0]['Component Group Name'] instead
        s_compfinal_group_tar.iloc[idx] = s_comp_grouplist[idx][1]
        s_compfinal_formula_tar.iloc[idx] = s_comp_grouplist[idx][2]
        s_compfinal_RT_tar.iloc[idx] = s_compfinal_RT.iloc[idx]
        s_masserr_tar.iloc[idx] = s_compfinal_masserr.iloc[idx]
        s_numhomolog_tar.iloc[idx] = s_numhomolog.iloc[idx]
        s_libscore_tar.iloc[idx] = s_compfinal_libscore.iloc[idx]
        s_IVconc_tar.iloc[idx, :] = s_IV_conc2.iloc[idx, :]
        s_compfinal_area_tar.iloc[idx, :] = s_compfinal_area_all2.iloc[idx, :]
        s_response_factor_tar.iloc[idx] = s_response_factor.iloc[idx]
        s_lowercal_tar.iloc[idx] = s_lowercal.iloc[idx]
        s_uppercal_tar.iloc[idx] = s_uppercal.iloc[idx]
        s_min_matrixconc_tar.iloc[idx] = s_matrix_min_conc.iloc[idx, :].min()
        s_max_matrixconc_tar.iloc[idx] = s_matrix_max_conc.iloc[idx, :].max()
    else:
        s_matrix_conc_LoQ_final.iloc[idx, :] = s_matrix_conc_LoQ.iloc[idx, :]
        s_matrix_sample_conc_final.iloc[idx, :] = s_matrix_sample_conc.iloc[idx, :]
        s_compfinal_final.iloc[idx] = comp[0]['Component Name']
        s_compfinal_calstd_final.iloc[idx] = comp[2]
        s_compfinal_mass_final.iloc[idx] = s_comp_grouplist[idx][3]
        s_compfinal_group_final.iloc[idx] = s_comp_grouplist[idx][1]
        s_compfinal_formula_final.iloc[idx] = s_comp_grouplist[idx][2]
        s_compfinal_RT_final.iloc[idx] = s_compfinal_RT.iloc[idx]
        s_masserr_final.iloc[idx] = s_compfinal_masserr.iloc[idx]
        s_numhomolog_final.iloc[idx] = s_numhomolog.iloc[idx]
        s_libscore_final.iloc[idx] = s_compfinal_libscore.iloc[idx]
        s_IVconc_final.iloc[idx, :] = s_IV_conc2.iloc[idx, :]
        s_compfinal_area_final.iloc[idx, :] = s_compfinal_area_all2.iloc[idx, :]
        s_response_factor_final.iloc[idx] = s_response_factor.iloc[idx]
        s_lowercal_final.iloc[idx] = s_lowercal.iloc[idx]
        s_uppercal_final.iloc[idx] = s_uppercal.iloc[idx]
        s_min_matrixconc_final.iloc[idx] = s_matrix_min_conc.iloc[idx, :].min()
        s_max_matrixconc_final.iloc[idx] = s_matrix_max_conc.iloc[idx, :].max()

### --- [intern] #16 on matlab script
padding = len(s_compfinal_final) + 3

# Write to tab1: SemiQuant Results
tab1 = pd.DataFrame(0, index=range(len(s_compfinal_final) + len(s_compfinal_tar) + 3), columns=range(len(s_unique_samples2) + 5), dtype=object)
tab1.iloc[0, 0] = 'Suspect Compound [' + units + ']'
tab1.iloc[1:len(s_compfinal_final) + 1, 0] = s_compfinal_final.values
tab1.iloc[0, 1] = 'Potential Isomer(s)'
tab1.iloc[1:len(s_compfinal_final) + 1, 1] = s_compfinal_mass_final.values
tab1.iloc[0, 2] = 'Formula'
tab1.iloc[1:len(s_compfinal_final) + 1, 2] = s_compfinal_formula_final.values
tab1.iloc[0, 3] = 'Precursor Mass'
tab1.iloc[1:len(s_compfinal_final) + 1, 3] = s_compfinal_mass_final.values
tab1.iloc[0, 4] = 'Group'
tab1.iloc[1:len(s_compfinal_final) + 1, 4] = s_compfinal_group_final.values
tab1.iloc[0, 5:len(s_unique_samples2) + 5] = s_unique_samples2
tab1.iloc[1:len(s_compfinal_final) + 1, 5:len(s_unique_samples2) + 5] = s_matrix_conc_LoQ_final.values
tab1[padding, 0] = 'Target Compound [' + units + ']'
tab1.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 0] = s_compfinal_tar.values
tab1.iloc[padding, 1] = 'Potential Isomer(s)'
tab1.iloc[padding + 1:padding + len(s_compfinal_mass_tar) + 1, 1] = s_compfinal_mass_tar.values
tab1.iloc[padding, 2] = 'Formula'
tab1.iloc[padding + 1:padding + len(s_compfinal_formula_tar) + 1, 2] = s_compfinal_formula_tar.values
tab1.iloc[padding, 3] = 'Precursor Mass'
tab1.iloc[padding + 1:padding + len(s_compfinal_mass_tar) + 1, 3] = s_compfinal_mass_tar.values
tab1.iloc[padding, 4] = 'Group'
tab1.iloc[padding + 1:padding + len(s_compfinal_group_tar) + 1, 4] = s_compfinal_group_tar.values
tab1.iloc[padding, 5:len(s_unique_samples2) + 5] = s_unique_samples2
tab1.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 5:len(s_unique_samples2) + 5] = s_matrix_conc_LoQ_tar.values

# tab2: suspect + target semi-quant matrix raw results (without dilution or RL)
tab2 = pd.DataFrame(0, index=range(len(s_compfinal_final) + len(s_compfinal_tar) + 3), columns=range(len(s_unique_samples2) + 3), dtype=object)
tab2.iloc[0, 0] = 'Suspect Compound [' + units + ']'
tab2.iloc[1:len(s_compfinal_final) + 1, 0] = s_compfinal_final.values
tab2.iloc[0, 1] = 'Precursor Mass'
tab2.iloc[1:len(s_compfinal_final) + 1, 1] = s_compfinal_mass_final.values
tab2.iloc[0, 2] = 'Group'
tab2.iloc[1:len(s_compfinal_group_final) + 1, 2] = s_compfinal_group_final.values
tab2.iloc[1:len(s_compfinal_final) + 1, 3:len(s_unique_samples2) + 3] = s_unique_samples2
tab2.iloc[1:len(s_compfinal_final) + 1, 3:len(s_unique_samples2) + 3] = s_matrix_sample_conc_final.values
tab2.iloc[padding, 0] = 'Target Compound [' + units + ']'
tab2.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 0] = s_compfinal_tar.values
tab2.iloc[padding, 1] = 'Precursor Mass'
tab2.iloc[padding + 1:padding + len(s_compfinal_mass_tar) + 1, 1] = s_compfinal_mass_tar.values
tab2.iloc[padding, 2] = 'Group'
tab2.iloc[padding + 1:padding + len(s_compfinal_group_tar) + 1, 2] = s_compfinal_group_tar.values
tab2.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 3:len(s_unique_samples2) + 3] = s_unique_samples2
tab2.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 3:len(s_unique_samples2) + 3] = s_matrix_sample_conc_tar.values

# tab3: suspect + target semi-quant in-vial concentration results
tab3 = pd.DataFrame(0, index=range(len(s_compfinal_final) + len(s_compfinal_tar) + 3), columns=range(len(s_unique_samples2) + 3), dtype=object)
tab3.iloc[0, 0] = 'Suspect Compound [ng/L] (no dilution)'
tab3.iloc[1:len(s_compfinal_final) + 1, 0] = s_compfinal_final.values
tab3.iloc[0, 1] = 'Precursor Mass'
tab3.iloc[1:len(s_compfinal_final) + 1, 1] = s_compfinal_mass_final.values
tab3.iloc[0, 2] = 'Group'
tab3.iloc[1:len(s_compfinal_group_final) + 1, 2] = s_compfinal_group_final.values
tab3.iloc[1:len(s_compfinal_final) + 1, 3:len(s_unique_samples2) + 3] = s_unique_samples2
tab3.iloc[1:len(s_compfinal_final) + 1, 3:len(s_unique_samples2) + 3] = s_IVconc_final.values
tab3.iloc[padding, 0] = 'Target Compound [ng/L] (no dilution)'
tab3.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 0] = s_compfinal_tar.values
tab3.iloc[padding, 1] = 'Precursor Mass'
tab3.iloc[padding + 1:padding + len(s_compfinal_mass_tar) + 1, 1] = s_compfinal_mass_tar.values
tab3.iloc[padding, 2] = 'Group'
tab3.iloc[padding + 1:padding + len(s_compfinal_group_tar) + 1, 2] = s_compfinal_group_tar.values
tab3.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 3:len(s_unique_samples2) + 3] = s_unique_samples2
tab3.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 3:len(s_unique_samples2) + 3] = s_IVconc_tar.values

# tab4: semiquant calculation
tab4 = pd.DataFrame(0, index=range(len(s_compfinal_final) + len(s_compfinal_tar) + 3), columns=range(9), dtype=object)
tab4.iloc[0, 0] = 'Suspect Compound'
tab4.iloc[1:len(s_compfinal_final) + 1, 0] = s_compfinal_final.values
tab4.iloc[0, 1] = 'Precursor Mass'
tab4.iloc[1:len(s_compfinal_final) + 1, 1] = s_compfinal_mass_final.values
tab4.iloc[0, 2] = 'Group'
tab4.iloc[1:len(s_compfinal_group_final) + 1, 2] = s_compfinal_group_final.values
tab4.iloc[0, 3] = 'Calibrant'
tab4.iloc[1:len(s_compfinal_final) + 1, 3] = s_compfinal_calibrant.values
tab4.iloc[0, 4] = 'Calibrant Response Factor'
tab4.iloc[1:len(s_compfinal_final) + 1, 4] = s_response_factor_final.values
tab4.iloc[0, 5] = 'Lower Cal Range'
tab4.iloc[1:len(s_compfinal_final) + 1, 5] = s_lowercal_final.values
tab4.iloc[0, 6] = 'Upper Cal Range'
tab4.iloc[1:len(s_compfinal_final) + 1, 6] = s_uppercal_final.values
tab4.iloc[0, 7] = 'Lower RL'
tab4.iloc[1:len(s_compfinal_final) + 1, 7] = s_min_matrixconc_final.values
tab4.iloc[0, 8] = 'Upper RL'
tab4.iloc[1:len(s_compfinal_final) + 1, 8] = s_max_matrixconc_final.values
tab4.iloc[padding, 0] = 'Target Compound'
tab4.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 0] = s_compfinal_tar.values
tab4.iloc[padding, 1] = 'Precursor Mass'
tab4.iloc[padding + 1:padding + len(s_compfinal_mass_tar) + 1, 1] = s_compfinal_mass_tar.values
tab4.iloc[padding, 2] = 'Group'
tab4.iloc[padding + 1:padding + len(s_compfinal_group_tar) + 1, 2] = s_compfinal_group_tar.values
tab4.iloc[padding, 3] = 'Calibrant'
tab4.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 3] = s_compfinal_calibrant_tar.values
tab4.iloc[padding, 4] = 'Calibrant Response Factor'
tab4.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 4] = s_response_factor_tar.values
tab4.iloc[padding, 5] = 'Lower Cal Range'
tab4.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 5] = s_lowercal_tar.values
tab4.iloc[padding, 6] = 'Upper Cal Range'
tab4.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 6] = s_uppercal_tar.values
tab4.iloc[padding, 7] = 'Lower RL'
tab4.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 7] = s_min_matrixconc_tar.values
tab4.iloc[padding, 8] = 'Upper RL'
tab4.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 8] = s_max_matrixconc_tar.values

# tab5: suspect + target raw peak areas
tab5 = pd.DataFrame(0, index=range(len(s_compfinal_final) + len(s_compfinal_tar) + 3), columns=range(len(s_unique_samples2) + 3), dtype=object)
tab5.iloc[0, 0] = 'Suspect Compound Peak Area'
tab5.iloc[1:len(s_compfinal_final) + 1, 0] = s_compfinal_final.values
tab5.iloc[0, 1] = 'Precursor Mass'
tab5.iloc[1:len(s_compfinal_final) + 1, 1] = s_compfinal_mass_final.values
tab5.iloc[0, 2] = 'Group'
tab5.iloc[1:len(s_compfinal_group_final) + 1, 2] = s_compfinal_group_final.values
tab5.iloc[1:len(s_compfinal_final) + 1, 3:len(s_unique_samples2) + 3] = s_unique_samples2
tab5.iloc[1:len(s_compfinal_final) + 1, 3:len(s_unique_samples2) + 3] = s_compfinal_area_final.values
tab5.iloc[padding, 0] = 'Target Compound [ng/L] (no dilution)' # TODO: copied from MATLAB script; is this correct?
tab5.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 0] = s_compfinal_tar.values
tab5.iloc[padding, 1] = 'Precursor Mass'
tab5.iloc[padding + 1:padding + len(s_compfinal_mass_tar) + 1, 1] = s_compfinal_mass_tar.values
tab5.iloc[padding, 2] = 'Group'
tab5.iloc[padding + 1:padding + len(s_compfinal_group_tar) + 1, 2] = s_compfinal_group_tar.values
tab5.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 3:len(s_unique_samples2) + 3] = s_unique_samples2
tab5.iloc[padding + 1:padding + len(s_compfinal_tar) + 1, 3:len(s_unique_samples2) + 3] = s_compfinal_area_tar.values

# tab6: confidence level assessment template
tab6 = pd.DataFrame(0, index=range(len(s_compfinal_final) + len(s_compfinal_tar) + 3), columns=range(10), dtype=object)
tab6.iloc[0, 0] = 'Suspect Confidence Level'
tab6.iloc[1:len(s_compfinal_final) + 1, 0] = s_compfinal_final.values
tab6.iloc[0, 1] = 'Precursor Mass'
tab6.iloc[1:len(s_compfinal_final) + 1, 1] = s_compfinal_mass_final.values
tab6.iloc[0, 2] = 'Group'
tab6.iloc[1:len(s_compfinal_group_final) + 1, 2] = s_compfinal_group_final.values
tab6.iloc[0, 3] = 'Formula'
tab6.iloc[1:len(s_compfinal_final) + 1, 3] = s_compfinal_formula_final.values
tab6.iloc[0, 4] = 'RT'
tab6.iloc[1:len(s_compfinal_final) + 1, 4] = s_compfinal_RT_final.values
tab6.iloc[0, 5] = 'Mass Error'
tab6.iloc[1:len(s_compfinal_final) + 1, 5] = s_masserr_final.values
tab6.iloc[0, 6] = 'Homologes Present'
tab6.iloc[1:len(s_compfinal_final) + 1, 6] = s_numhomolog_final.values
tab6.iloc[0, 7] = 'Library Score'
tab6.iloc[1:len(s_compfinal_final) + 1, 7] = s_libscore_final.values
tab6.iloc[0, 8] = 'Confidence Level'
# TODO: the filler for this information was commented out in the MATLAB.
tab6.iloc[0, 9] = 'Fragments Annotated'
# TODO: the filler for this information was commented out in the MATLAB.

# Write all the tabs to an Excel file
# TODO: fix this, work on this
output_pathname = os.path.join(suspect_dirname, f"{suspect_filename}_results.xlsx")
with pd.ExcelWriter(output_pathname, engine='openpyxl') as writer:
    tab1.to_excel(writer, sheet_name='SemiQuant Results', index=False, header=False)
    tab2.to_excel(writer, sheet_name='Raw Matrix Conc', index=False, header=False)
    tab3.to_excel(writer, sheet_name='Raw InVial Conc', index=False, header=False)
    tab4.to_excel(writer, sheet_name='SemiQuant Calculation', index=False, header=False)
    tab5.to_excel(writer, sheet_name='Raw Peak Area', index=False, header=False)
    tab6.to_excel(writer, sheet_name='Confidence Level Assessment', index=False, header=False)

print(f">> Done. Results written to {output_pathname}")