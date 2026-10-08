#!/usr/bin/env python3

##############################################
#
#   target-data processing rewrite
#   Last modified: 2026-07-10
#
##############################################
# Ensure you've installed all the dependencies from requirements.txt before proceeding.
# A rewrite of the original MATLAB script (Conrad Pritchard, v5.2.1, 2/12/2026) using pandas; rewritten by Shubhang Vyas (2026-07-10).

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
samplemass_path = ""
# See if the user wants to specify a sample mass path.
if input("Do you want to specify a sample mass file?").lower() in ["y", "yes"]:
    samplemass_filename = input("Enter the name of the Excel file to set as the sample mass file: ")
    # If the user gave a bad file extension that is not .xlsx or .xls, notify the user and exit with error code 1 (FAILURE).
    if re.search(r'\.(?!xlsx|xls$)', samplemass_filename):
        print(">>> Not a valid file extension for the sample mass file -- must be .xlsx or .xls. <<<")
        exit(1)
    
    # If there was no extension, add .xlsx to the end of the file name.
    if not re.search(r'\.(xlsx|xls)$', samplemass_filename):
        samplemass_filename += ".xlsx"
    
    # Get the directory folder name that has this file.
    samplemass_dirname = input("Enter the path of the directory with the sample mass file (or leave blank for current directory): ")
    # If the user didn't provide a path, use the directory of the target file.
    if not samplemass_dirname:
        print("> Using the same directory as the target file for the sample mass file.")
        samplemass_dirname = path_name
    
    # Make the full path for the sample mass file.
    samplemass_path = os.path.join(samplemass_dirname, samplemass_filename)
    
    # If the path doesn't exist, notify the user and exit with error code 1 (FAILURE).
    if not os.path.exists(samplemass_path):
        print(f">>> Couldn't find a sample mass file at {samplemass_path}! <<<")
        exit(1)

##### 3) Get options for including method blanks and injection type.
# Determine if we want to subtract method blanks from the samples.
# Keep asking the user until they give a valid response.
subtract_mb = False
while True:
    user_asked_subtract_mb = input("Subtract method blanks from the samples? (y/n/yes/no): ")
    if user_asked_subtract_mb.lower() in ['y', 'yes']:
        sub_mb = True
        break
    if user_asked_subtract_mb.lower() in ['n', 'no']:
        sub_mb = False
        break

# Determine if the user used `direct injection` in the samples.
# Keep asking the user until they give a valid response.
direct_injection = False
while True:
    user_asked_direct_injection = input("Was direct injection used? (if not, SPE will be assumed) -- (y/n/yes/no): ")
    if user_asked_direct_injection.lower() in ['y', 'yes']:
        direct_injection = True
        break
    if user_asked_direct_injection.lower() in ['n', 'no']:
        direct_injection = False
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

# Make sure that the required columns that we need, exist.
columns_to_check = [
    "Sample Name", 
    "Sample Index", 
    "Injection Volume", 
    "Component Name", 
    "Component Group Name", 
    "Component Type", 
    "Retention Time", 
    "Precursor Mass", 
    "Mass Error Confidence", 
    "Mass Error (ppm)", 
    "IS Name", 
    "Area", 
    "IS Area", 
    "Actual Concentration", 
    "Calculated Concentration", 
    "Sample Type", 
    "Used", 
    "Polarity"
]
for col in columns_to_check:
    # If the column doesn't exist, notify the user and exit with error code 1 (FAILURE).
    if col not in df.columns:
        print(f">>> Missing required column in data: {col} <<<")
        exit(1)

#########################################################
### -- Determine analysis method and set constants -- ###
#########################################################
# Select the injection volume: first row (iloc[0]), column `Injection Volume`; then, convert it from a string into an integer.
injection_volume = int(df.iloc[0]["Injection Volume"])
# Select the polarity: first row (iloc[0]), column `Polarity`. 
polarity = df.iloc[0]["Polarity"]

# TODO: move this
matrix_extraction_factors = {
    # (total vial volume / extract vial volume) * total extraction volume / 1000 [mL/L]; reports into units [L]
    "soil"      : (0.4 / 0.1)  * (1.5 / 1000),
    "dust"      : (0.6 / 0.45) * (1.5 / 1000) * (5 / 0.25),
    "water"     : (0.75 / 0.375) * (5 / 1000) * (1000),
    "SPEsoil"   : (0.75 / 0.375) * (5 / 1000) * (32 / 2.5),
}
# Aliases
matrix_extraction_factors["SPEwater"] = matrix_extraction_factors["water"]


# We need to determine these values, so we'll set them to default values for now.
method_type = "UNKNOWN"
CCV_actual_concentration = 0.0  # CCV/QC concentration [ng/L]
ISC_actual_concentration = 0.0  # ISC concentration [ng/L]
EPA_actual_concentration = 0.0  # EPA Bullseye concentration [ng/L]

# If the injection volume was 100 uL:
if injection_volume == 100:    
    if direct_injection:
        # 100ul volume with direct injection (DI).
        method_type = "100uL-DI"
        CCV_actual_concentration = 200.0
        ISC_actual_concentration = 33.33
        EPA_actual_concentration = 333.33
    elif polarity == "Positive":
        # 100ul volume, NO direct injection, positive polarity (via SPE method).
        method_type = "100uL-SPE-Positive"
        CCV_actual_concentration = 200.0
        ISC_actual_concentration = 10.0
        # Leave EPA_actual_concentration as 0.0: not needed in this method.
    elif polarity == "Negative":
        # 100ul volume, NO direct injection, negative polarity (via SPE method).
        method_type = "100uL-SPE-Negative"
        CCV_actual_concentration = 250.0
        ISC_actual_concentration = 25.0
        EPA_actual_concentration = 250.0
elif injection_volume == 1000:
    # Always 1000uL direct injection -- aqueous method.
    # 1000uL = 1mL, but leave at 1000uL for readability's sake.
    method_type = "1000uL-DI"
    CCV_actual_concentration = 200.0
    ISC_actual_concentration = 33.33
    EPA_actual_concentration = 333.33

# If we haven't set our method type yet, notify the user and exit with error code 1 (FAILURE).
if method_type == "UNKNOWN":
    print(f">>> Unknown method type for injection volume {injection_volume} uL and polarity {polarity}! (di: {direct_injection}) <<<")
    exit(1)

#########################################################
### Identify/categorize the components and samples -- ###
#########################################################

# Get all columns with the same sample name as the first row.
first_sample_name = df.iloc[0]["Sample Name"]
of_first_sample = df[df["Sample Name"] == first_sample_name]

# ...DEBUG
print(of_first_sample["Component Type"].tolist())

# Get all samples that are quantifiers/qualifiers, but NOT injection standards.
targets = of_first_sample[
    (of_first_sample["Component Type"].isin(["Quantifiers", "Qualifiers"])) &
    (of_first_sample["Component Group Name"] != "Inj Stds")
]

# ...DEBUG
print(len(targets), targets)

# Of our `targets`, collect the component names, masses, groups, and their internal standard names.
component_names = targets["Component Name"].tolist()
component_masses = targets["Precursor Mass"].tolist()
component_groups = targets["Component Group Name"].tolist()
component_IntS_names = targets["IS Name"].tolist()

# ... DEBUG [TODO] found unused
num_components = len(component_names)

# If we're using 100uL, find our injection standards and collect their names.
if injection_volume == 100:
    injection_standards = of_first_sample[
        of_first_sample["Component Group Name"] == "Inj Stds"
    ]

    InjS_names = injection_standards["Component Name"].tolist()
    num_InjS = len(InjS_names) # ...DEBUG [TODO] found unused
else:
    InjS_names = []
    num_InjS = 0 # ...DEBUG [TODO] found unused

# Collect all internal standards (that are also NOT injection standards), and count.
# ...DEBUG [TODO] found unused
num_internal_standards = len(
    of_first_sample[
        (of_first_sample["Component Type"].isin(["Internal Standard", "Internal Standards"])) &
        (of_first_sample["Component Group Name"] != "Inj Stds")
    ]
)

# Sum the total number of components and internal standards
# ...DEBUG [TODO] found unused
total_num_components = num_components + num_internal_standards

##########################
### -- Extract data -- ###
##########################

# Note that this script, unlike the original MATLAB script, does not manually keep counters and iterate over the rows.
# This segment, however, is essentially identical to the corresponding MATLAB section.

# A more in-depth explanation as to why this has an identical approach to the MATLAB:
# - The MATLAB maintains 4 counters: ncal (number of calibration standards), naQC (number of analytical QC samples),
#   nmQC (number of method QC samples), and nsam (number of samples).
# - Those counters are incremented when the script encounters a *new sample* -- when it does, however,
#   it only increments the counter of the current component type it's on.
# 
# A visual explanation:
#       id          sample [determined by script]           component type
#       1           1                                       aQC                 <-- ncal=0; **naQC=1;** nmQC=0; nsam=0
#       2           1                                       aQC
#       3           1                                       aQC
#       4           2                                       mQC                 <-- ncal=0; naQC=1; **nmQC=1;** nsam=0
#       5           2                                       mQC
#       6           2                                       mQC
#       7           3                                       sample              <-- ncal=0; naQC=1; nmQC=1; **nsam=1**
#       8           3                                       sample
#       9           3                                       sample
#       10          4                                       mQC                 <-- ncal=0; naQC=1; **nmQC=2;** nsam=1
#       11          4                                       mQC
#       12          4                                       mQC
# 
# - The pandas script does the same thing internally - before building its pivot tables (same thing as the dataframes
#   being constructed in the MATLAB script), it first differentiates by type:
# 
#   aQC samples:
#       1           1                                       aQC
#       2           1                                       aQC
#       3           1                                       aQC
#   
#   mQC samples:
#       4           2                                       mQC
#       5           2                                       mQC
#       6           2                                       mQC
#       10          4                                       mQC
#       11          4                                       mQC
#       12          4                                       mQC
#   [... other component types]
# 
# - Note: after differentiating by component type, when iterating through and detecting new samples (via the sample id in our abstraction),
#   we now maintain individual counters per component type which our pivot tables respect.

# As for differentiating between samples, use a loose hack from the original MATLAB script: simply check if the component name equals the component name of the first row.
# This is, of course, under the assumptions that:
# 1) component names are *never* repeated in the same sample, and
# 2) the first row of every sample is always the same component name.
# 
# Of course, this is a loose hack, but under the assumption that all component names are unique in the same sample,
# this should work fine.
# 
# example:
# - id1     ComponentA  <-- sample 1
# - id2     ComponentB
# - id3     ComponentC
# - id4     ComponentD
# - id5     ComponentA  <-- name repeated, new sample
# - id6     ComponentB
# - id7     ComponentC
# - id8     ComponentD
# - id9     ComponentA  <-- name repeated, new sample
# ...

def build_pivot(df, value_column, index_list=component_names, to_numeric=True):
    # First, obtain a unique (no repeat) order of our sample indices.
    col_order = df["Sample Index"].unique().tolist()

    # Convert our 2D dataframe into a pivot table.
    table = df.pivot_table(
        index="Component Name",
        columns="Sample Index",
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

### Define our df filters that we'll use to distinguish between different component types.
# Filter for calibration standards: if the sample type **equals** "Standard"
is_calstd = df["Sample Type"] == "Standard"
# Filter for analytical QC samples: if the sample type **is not** a calibration standard, *and* the sample name
# **contains** either "CCV", "ISC", "LB", "EPA", "ICV", or "QC".
is_aQC = ~is_calstd & re.search(r'CCV|ISC|LB|EPA|ICV|QC', df["Sample Name"])
# Filter for method QC samples: if the sample type **is not** a calibration standard or an analytical QC, *and*
# the sample name **contains** either "LCS", "LLLCS", "LCSD", or "MB".
#   Note that we only match for "MB" and "LCS", because "LCS", "LLLCS", and "LCSD" all contain the pattern "LCS".
is_mQC = ~is_calstd & ~is_aQC & re.search(r'LCS|MB', df["Sample Name"])
# Filter for "skips" (entries to ignore): if the sample type **is not** a calibration standard, analytical QC,
# or method QC, *and* the sample name **contains** either "DB", "blank check", or "AFFF".
#   Note that this filter is not actually used to extract any rows from our df, but serves useful for readability
#   when extracting samples.
is_skip = ~is_calstd & ~is_aQC & ~is_mQC & re.search(r'DB|blank check|AFFF', df["Sample Name"], re.IGNORECASE)
# Filter for samples: if the sample type **is not** a calibration standard, analytical QC, method QC, or a skip.
is_sample = ~is_calstd & ~is_aQC & ~is_mQC & ~is_skip

### Define our df filters to distinguish between targets, injection standards, or internal standards.
# Filter for targets: if the component type **is** either "Quantifiers" or "Qualifiers", *and*
# the component group name **does not equal** "Inj Stds".
is_target = (
    (df["Component Type"].isin(["Quantifiers", "Qualifiers"])) &
    (df["Component Group Name"] != "Inj Stds")
)
# Filter for injection standards: if the component group name **equals** "Inj Stds".
#   Note that this script only extracts injection standards if this is a 100uL-SPE method.
is_InjS = df["Component Group Name"] == "Inj Stds"

# Extract our targets and injection standards.

calstd_targets = df[is_calstd & is_target]
aQC_targets = df[is_aQC & is_target]
mQC_targets = df[is_mQC & is_target]
sample_targets = df[is_sample & is_target]

calstd_InjS = df[is_calstd & is_InjS]
aQC_InjS = df[is_aQC & is_InjS]
mQC_InjS = df[is_mQC & is_InjS]
sample_InjS = df[is_sample & is_InjS]

calstd_injPA         = build_pivot(calstd_InjS, "Area", index_list=InjS_names)
calstd_realIS        = build_pivot(calstd_targets, "IS Actual Concentration")
calstd_realconc      = build_pivot(calstd_targets, "Actual Concentration")
calstd_IntSPA        = build_pivot(calstd_targets, "IS Area")
calstd_compconc      = build_pivot(calstd_targets, "Calculated Concentration")
calstd_PA            = build_pivot(calstd_targets, "Area")
calstd_sidx          = build_pivot(calstd_targets, "Sample Index", to_numeric=False)
calstd_used          = build_pivot(calstd_targets, "Used", to_numeric=False)

aQC_injPA            = build_pivot(aQC_InjS, "Area", index_list=InjS_names)
aQC_realIS           = build_pivot(aQC_targets, "IS Actual Concentration")
aQC_IntSPA           = build_pivot(aQC_targets, "IS Area")
aQC_compconc         = build_pivot(aQC_targets, "Calculated Concentration")
aQC_PA               = build_pivot(aQC_targets, "Area")
aQC_sidx             = build_pivot(aQC_targets, "Sample Index", to_numeric=False)
aQC_names            = build_pivot(aQC_targets, "Component Name", to_numeric=False)

mQC_injPA            = build_pivot(mQC_InjS, "Area", index_list=InjS_names)
mQC_realIS           = build_pivot(mQC_targets, "IS Actual Concentration")
mQC_IntSPA           = build_pivot(mQC_targets, "IS Area")
mQC_compconc         = build_pivot(mQC_targets, "Calculated Concentration")
mQC_PA               = build_pivot(mQC_targets, "Area")
mQC_sidx             = build_pivot(mQC_targets, "Sample Index", to_numeric=False)
mQC_names            = build_pivot(mQC_targets, "Component Name", to_numeric=False)

samples_injPA        = build_pivot(sample_InjS, "Area", index_list=InjS_names)
samples_realIS       = build_pivot(sample_targets, "IS Actual Concentration")
samples_IntSPA       = build_pivot(sample_targets, "IS Area")
samples_compconc     = build_pivot(sample_targets, "Calculated Concentration")
samples_PA           = build_pivot(sample_targets, "Area")
samples_sidx         = build_pivot(sample_targets, "Sample Index", to_numeric=False)

# TODO: remove aliases
# ...DEBUG -- quick aliases to make the script work
calstd_injPA_tble = calstd_injPA
calstd_realIS_tble = calstd_realIS
calstd_realconc_tble = calstd_realconc
calstd_IntSPA_tble = calstd_IntSPA
calstd_compconc_tble = calstd_compconc
calstd_PA_tble = calstd_PA
calstd_sidx_tble = calstd_sidx
calstd_used = calstd_used

aQC_injPA_tble = aQC_injPA
aQC_realIS_tble = aQC_realIS
aQC_IntSPA_tble = aQC_IntSPA
aQC_compconc_tble = aQC_compconc
aQC_PA_tble = aQC_PA
aQC_sidx_tble = aQC_sidx
aQC_names_tble = aQC_names

mQC_injPA_tble = mQC_injPA
mQC_realIS_tble = mQC_realIS
mQC_IntSPA_tble = mQC_IntSPA
mQC_compconc_tble = mQC_compconc
mQC_PA_tble = mQC_PA
mQC_sidx_tble = mQC_sidx
mQC_names_tble = mQC_names

samples_injPA_tble = samples_injPA
samples_realIS_tble = samples_realIS
samples_IntSPA_tble = samples_IntSPA
samples_compconc_tble = samples_compconc
samples_PA_tble = samples_PA
samples_sidx_tble = samples_sidx

#########################################
### -- Determine calibration range -- ###
#########################################
# matlab orig: [calsort, calorderindex] = sort(calactconc(1, :), 2);
# TODO: is this right?
sort_order = calstd_realconc_tble.loc[0, :].apply(pd.to_numeric, errors='coerce').argsort()
calstd_realconc_sorted = calstd_realconc_tble.iloc[:, sort_order]
calstd_compconc_sorted = calstd_compconc_tble.iloc[:, sort_order]
calstd_used_sorted = calstd_used.iloc[:, sort_order]

# Lower/upper limits for each component
lowlim = {}
highlim = {}

# For each component:
for c in component_names:
    valid = []

    # For each calibration level (that was used),
    # if our error bound of the concentration was within (0.7, 1.3), append it to our list of valid concentrations.
    for calstd in calstd_realconc_sorted.columns:
        # The real and computed concentrations for this calibration standard and component.
        realconc = calstd_realconc_sorted.loc[calstd, c]
        compconc = calstd_compconc_sorted.loc[calstd, c]
        # Whether the calibration standard was used for this component.
        used = calstd_used_sorted.loc[calstd, c]
        acc = compconc / realconc if realconc != 0 else np.nan
        if 0.7 < acc < 1.3 and used.lower() == "true":
            valid.append(realconc)

    # TODO: probably not right?
    # lowlim[c] = valid[0] if valid else np.nan
    # highlim[c] = valid[-1] if valid else np.nan

    # but it's already sorted...
    # lowlim[c] = min(valid) if len(valid) > 0 else np.nan
    # highlim[c] = max(valid) if len(valid) > 0 else np.nan

    lowlim[c] = valid[0] if len(valid) > 0 else np.nan
    highlim[c] = valid[-1] if len(valid) > 0 else np.nan

#################################################################
### -- Determine peak area recovery for internal standards -- ###
#################################################################
if method_type == "100ul-SPE-Negative":
    # Assign injection standards for each component -- read the sample mass path.
    if not samplemass_path:
        # Notify the user and exit with error code 1 (FAILURE).
        print(">>> No sample mass file specified (found a 100uL-SPE-Negative method)! <<<")
        exit(1)
    InjS_list = pd.read_excel(
        samplemass_path,
        sheet_name="NIS",
        header=0,
        dtype=str
    )
    # Associates each component name with its corresponding injection standard name,
    # `{ component: injection_standard, ... }`
    # for an O(1) lookup.
    InjS_dict = dict(zip(InjS_list.iloc[:, 0], InjS_list.iloc[:, 1]))
    
    calstd_InjS_matched_PA = []
    aQC_InjS_matched_PA = []
    mQC_InjS_matched_PA = []
    samples_InjS_matched_PA = []

    # Build a matrix of injection standard peak areas (for each calstd, aQC, mQC, and sample)
    used_InjS_list = []
    missing = []
    for c in component_names:
        # If we have an injection standard for this component:
        if c in InjS_dict:
            # Get the injection standard used for this component.
            InjS = InjS_dict[c]
            # Find the index of the injection standard in our previously computed list.
            InjS_idx = InjS_names.index(InjS)
            
            # Add this injection standard to our list of used injection standards.
            used_InjS_list.append(InjS)

            # Save the peak areas of this injection standard from each component type (calstd, aQC, mQC, sample).
            calstd_InjS_matched_PA.append(calstd_injPA[injS_idx])
            # TODO: the conditional check `num_aQC > 0` was not in the original MATLAB script and is not featured in the ipynb rewrite. Remove?
            # TODO: num_aQC isn't a thing
            if num_aQC > 0:
                aQC_InjS_matched_PA.append(aQC_injPA[injS_idx])
            if len(mQC_compconc_tble) > 0:
                mQC_InjS_matched_PA.append(mQC_injPA[injS_idx])
            samples_InjS_matched_PA.append(samples_injPA[injS_idx])
        else:
            missing.append(c)
    
    # If we're missing injection standards for any component, notify the user and exit with error code 1 (FAILURE).
    if len(missing) > 0:
        print(f">>> Missing injection standards for the following components: {', '.join(missing)} <<<")
        exit(1)
    
    # TODO: remove this matlab rough write
    samples_IS_recovery = samples_IntSPA / samples_InjS_matched_PA / mean(calstd_IntSPA / calstd_InjS_matched_PA) * mean(calstd_realIS) / samples_realIS

    # In another rewrite:
    calstd_InjS_matched_PA = pd.DataFrame(calstd_InjS_matched_PA, index=component_names).T
    aQC_InjS_matched_PA = pd.DataFrame(aQC_InjS_matched_PA, index=component_names).T
    mQC_InjS_matched_PA = pd.DataFrame(mQC_InjS_matched_PA, index=component_names).T
    samples_InjS_matched_PA = pd.DataFrame(samples_InjS_matched_PA, index=component_names).T

    # Compute the mean calibration-standard injection peak area to use for recovery normalization.
    calstd_realIS_mean = calstd_realIS_tble.mean(axis=1)
    calstd_injPA_mean = calstd_injPA_tble.mean(axis=1).values[:, None]

    # Find the **internal** standard recovery for each component (for each type: calstd, aQC, mQC, sample).
    common_ratio = (calstd_IntSPA / calstd_InjS_matched_PA).mean(axis=1).values[:, None]
    # common_ratio = (calstd_IntSPA_mean / calstd_injS_mean).values[:, None]

    samples_IntS_recovery = (samples_IntSPA / samples_InjS_matched_PA)  / common_ratio * (calstd_realIS_mean / samples_realIS)
    aQC_IntS_recovery     = (aQC_IntSPA / aQC_InjS_matched_PA)          / common_ratio * (calstd_realIS_mean / aQC_realIS)
    mQC_IntS_recovery     = (mQC_IntSPA / mQC_InjS_matched_PA)          / common_ratio * (calstd_realIS_mean / mQC_realIS)

    # Find the **injection** standard recovery for each component (for each type: calstd, aQC, mQC, sample).
    samples_InjS_recovery = samples_injPA   / calstd_injPA_mean
    aQC_InjS_recovery     = aQC_injPA       / calstd_injPA_mean
    mQC_InjS_recovery     = mQC_injPA       / calstd_injPA_mean

else:
    # Compute the internal standard recovery for each sample type, taking into account any differences between real internal standard concentrations (using our calibration standard
    # as the reference point for the expected internal standard peak area)
    calstd_IntSPA_mean = calstd_IntSPA.mean(axis=1).values[:, None]
    calstd_realIS_mean = calstd_realIS.mean(axis=1)

    samples_IntS_recovery = (samples_IntSPA / calstd_IntSPA_mean)      * (calstd_realIS_mean / samples_realIS)
    aQC_IntS_recovery = (aQC_IntSPA / calstd_IntSPA_mean)              * (calstd_realIS_mean / aQC_realIS)
    mQC_IntS_recovery = (mQC_IntSPA / calstd_IntSPA_mean)              * (calstd_realIS_mean / mQC_realIS)

############################################
### -- Check LCMS analytical accuracy -- ###
############################################

# Find recoveries for CCV/QC/ISC/LB/EPA/ICV samples in our aQC samples.
aQC_columns = aQC_compconc_tble.columns

aQC_CCV_cols = [col for col in aQC_columns if "CCV" in col]
aQC_QC_cols  = [col for col in aQC_columns if "QC" in col]
aQC_ISC_cols = [col for col in aQC_columns if "ISC" in col]
aQC_LB_cols  = [col for col in aQC_columns if "LB" in col]
aQC_EPA_cols = [col for col in aQC_columns if "EPA" in col or "ICV" in col]

aQC_CCV_recovery = aQC_compconc_tble[aQC_CCV_cols] / CCV_actual_concentration       if len(aQC_CCV_cols) > 0 else pd.DataFrame()
aQC_QC_recovery  = aQC_compconc_tble[aQC_QC_cols]  / CCV_actual_concentration       if len(aQC_QC_cols) > 0 else pd.DataFrame()
aQC_ISC_recovery = aQC_compconc_tble[aQC_ISC_cols] / ISC_actual_concentration       if len(aQC_ISC_cols) > 0 else pd.DataFrame()
aQC_LB_recovery  = aQC_compconc_tble[aQC_LB_cols]                                   if len(aQC_LB_cols) > 0 else pd.DataFrame()
aQC_EPA_recovery = aQC_compconc_tble[aQC_EPA_cols] / EPA_actual_concentration       if len(aQC_EPA_cols) > 0 else pd.DataFrame()

aQC_all = pd.concat([aQC_CCV_recovery, aQC_QC_recovery, aQC_ISC_recovery, aQC_LB_recovery, aQC_EPA_recovery], axis=1)
aQC_all_names = aQC_CCV_cols + aQC_QC_cols + aQC_ISC_cols + aQC_LB_cols + aQC_EPA_cols

#################################################################
### -- Check method accuracy and determine reporting limit -- ###
#################################################################

# Reporting limit.
RL = pd.Series(lowlim, index=component_names, dtype=float)

if len(mQC_compconc_tble) > 0:
    # Calculate LCS+LCSD+LLLCSs and MBs in vial concentration (ng/L) for each component.
    mQC_LCS_cols = [c for c in mQC_compconc_tble.columns if "LCS" in c]
    mQC_MB_cols  = [c for c in mQC_compconc_tble.columns if "MB" in c]
    
    LCS_conc = mQC_compconc_tble[mQC_LCS_cols] if len(mQC_LCS_cols) > 0 else pd.DataFrame()
    MB_conc  = mQC_compconc_tble[mQC_MB_cols] if len(mQC_MB_cols) > 0 else pd.DataFrame()
    
    if mQC_MB_cols:
        # Compute the mean method blank concentration.
        # If we're subtracting method blanks (subtract_mb=True), subtract the mean from our values.
        MB_mean = MB_conc.mean(axis=1)

        if subtract_mb:
            # Subtract the mean method blank concentration from the LCS concentration.
            uncorrected_samples_compconc_tble = samples_compconc_tble.copy()
            samples_compconc_tble = samples_compconc_tble.sub(MB_mean, axis=0)

            # Subtract the mean method blank concentration from the MB concentrations.
            MB_conc = MB_conc.sub(MB_mean, axis=0)
        
        # If there are LCS columns as well:
        if mQC_LCS_cols:
            # Subtract the mean method blank concentration from the LCS concentrations.
            LCS_conc = LCS_conc.sub(MB_mean, axis=0)

        # Sets the reporting limit (RL) to the highest MB concentration or `lowlim`, whichever is greater.
        for c in component_names:
            # x3 -- used from the original MATLAB script (Conrad Pritchard, v5.2.1, 2/12/2026).
            MB_max = MB_conc[c].max() * 3 if c in MB_conc.columns else 0
            if subtract_mb:
                MB_max -= MB_mean[c] if c in MB_mean.index else 0
            
            RL[c] = max(lowlim[c], MB_max)    
    
    subtracted_MB_values = MB_mean if subtract_mb and aQC_MB_cols else pd.Series(0, index=component_names, dtype=float)
    mQC_all = pd.concat([LCS_conc, MB_conc], axis=1)
    mQC_all_names = mQC_LCS_cols + mQC_MB_cols

else:
    mQC_all = pd.DataFrame()
    mQC_all_names = []
    subtracted_MB_values = pd.Series(0, index=component_names, dtype=float)

############################################################################################################
### -- Calculate matrix concentration, matrix LoQ range, and replace values outside of our matrix LoQ -- ###
############################################################################################################

# If this is a 100uL SPE method:
if method_type in ["100uL-SPE-Positive", "100uL-SPE-Negative"]:
    # Read in injection standard matches -- our soil data.
    soildata = pd.read_excel(
        samplemass_path,
        sheet_name=test_name, # The script assumes that the name of the sheet with our soil data is the same as the test name.
        header=0,
        dtype=str
    )
    
    sample_names = soildata.iloc[:, 0]
    wet_mass = pd.to_numeric(soildata.iloc[:, 1], errors='coerce').tolist()
    moisture_frac = pd.to_numeric(soildata.iloc[:, 2], errors='coerce').tolist()
    matrix_type = soildata.iloc[:, 3].tolist()

    dry_mass = wet_mass * (1 - moisture_frac)

    # TODO: is sample_compconc_tble.columns the same thing as sample_names or samples_names_tble.index?
    matrix_sample_conc = pd.DataFrame(index=component_names, columns=sample_compconc_tble.columns, dtype=float)
    matrix_highlim = pd.DataFrame(index=component_names, columns=sample_compconc_tble.columns, dtype=float)
    matrix_lowlim = pd.DataFrame(index=component_names, columns=sample_compconc_tble.columns, dtype=float)

    for i in range(num_samples):
        # Find the index of this sample in soildata.
        sample = samples_names_tble.index[i]
        # TODO: is this even right?
        soilrow_match = soildata[sample_names == sample]
        if soilrow_match.empty:
            print(f">>> Sample {sample} not found in soil data! <<<")
            exit(1)
        
        soil_idx = soilrow_match.index[0]
        mtype = matrix_type[soil_idx].strip()

        if mtype not in matrix_extraction_factors:
            print(f">>> Unknown matrix type/extraction factor {mtype} for sample {sample}! <<<")
            exit(1)

        extraction_factor = matrix_extraction_factors[mtype]
        mass = dry_mass[soil_idx]
        common_factor = extraction_factor / mass
        
        matrix_sample_conc.iloc[:, i] = samples_compconc_tble.iloc[:, i]            * common_factor
        matrix_highlim.iloc[:, i]     = pd.Series(highlim, index=component_names)    * common_factor
        matrix_lowlim.iloc[:, i]      = RL                                          * common_factor

    # Water: ng/L; solids: ng/g
    units = "ng/L" if matrix_type.iloc[-1].strip() == "water" else "ng/g"
else:
    # Multiply the sample computed concentration by vial volume (1.5mL) / volume water sample (0.9mL)
    matrix_sample_conc = sample_compconc_tble * (1.5 / 0.87)
    matrix_highlim = pd.DataFrame(
        { col: pd.Series(highlim, index=component_names) * factor for col in sample_compconc_tble.columns }
    )
    matrix_lowlim = pd.DataFrame(
        { col: RL * factor for col in sample_compconc_tble.columns }
    )
    units = "ng/L"


# Values outside of our LoQ will get replaced with comparisons, since we can't report values outside of our LoQ.
# matrix_sample_compconc_LoQ: the *LoQ* of our *matrix* of *computed concentrations* for our *samples*
# outside_matrix_identifiers: if there are values that are NaN or outside of our LoQ, we want to indicate if they're greater/less than a certain number. Instead of specifying
#   the number right now, we instead set the number in matrix_sample_compconc_LoQ, and leave a character indicator in outside_matrix_identifiers (such as "<" or ">").
#   If the character value at a specific row/column in outside_matrix identifiers is '' (empty), that corresponding value in matrix_sample_compconc_LoQ is valid and can be reported as-is.
matrix_sample_compconc_LoQ = matrix_sample_conc.copy().astype(object)
outside_matrix_identifiers = pd.DataFrame("", index=component_names, columns=matrix_sample_conc.columns, dtype=str)
for c in component_names:
    for s in matrix_sample_conc.columns:
        value = matrix_sample_conc.loc[c, s]
        high = matrix_highlim.loc[c, s]
        low = matrix_lowlim.loc[c, s]

        if pd.isna(value) or value < low:
            matrix_sample_compconc_LoQ.loc[c, s] = low
            outside_matrix_identifiers.loc[c, s] = "<"
        elif value > high:
            matrix_sample_compconc_LoQ.loc[c, s] = high
            outside_matrix_identifiers.loc[c, s] = ">"
        else:
            matrix_sample_compconc_LoQ.loc[c, s] = value
            outside_matrix_identifiers.loc[c, s] = ""

####################################
### -- Correct diluted values -- ###
####################################

# Same rules as above, but for diluted values.
diluted_matrix_sample_compconc_LoQ = matrix_sample_compconc_LoQ.copy()
diluted_outside_matrix_identifiers = outside_matrix_identifiers.copy()

for s in samples_compconc_tble.columns:
    diluted_matcher = re.search(r"_d(\d+)x", s, re.IGNORECASE)
    if diluted_matcher:
        dilution_factor = int(diluted_matcher.group(1))
        diluted_matrix_sample_compconc_LoQ[s] = matrix_sample_compconc_LoQ[s] * dilution_factor

matrix_sample_compconc_LoQ_final = diluted_matrix_sample_compconc_LoQ.copy()
for c in component_names:
    for s in diluted_matrix_sample_compconc_LoQ.columns:
        flag = diluted_outside_matrix_identifiers.loc[c, s]
        value = diluted_matrix_sample_compconc_LoQ.loc[c, s]

        if flag == "<":
            matrix_sample_compconc_LoQ_final.loc[c, s] = f"<{value:.2f}"
        elif flag == ">":
            matrix_sample_compconc_LoQ_final.loc[c, s] = f">{value:.2f}"
        else:
            matrix_sample_compconc_LoQ_final.loc[c, s] = f"{value:.2f}"

#############################################
### -- Write all data to an Excel file -- ###
#############################################

out_path = os.path.join(out_dir, f"{test_name}_results.xlsx")
with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
    matrix_sample_compconc_LoQ_final.to_excel(writer, sheet_name="Sample Concentrations", index=True)
    samples_IS_recovery.to_excel(writer, sheet_name="Surrogate Recovery", index=True)
    if len(mQC_all) > 0:
        mQC_all.to_excel(writer, sheet_name="Method Accuracy", index=True)
    else:
        pd.DataFrame(["*** NO METHOD ACCURACY SAMPLES FOUND! ***"]).to_excel(
            writer, sheet_name="Method Accuracy", index=False, header=False
        )
    
    if len(aQC_all) > 0:
        aQC_all.to_excel(writer, sheet_name="Analytical Accuracy", index=True)
    else:
        pd.DataFrame(["*** NO ANALYTICAL ACCURACY SAMPLES FOUND! ***"]).to_excel(
            writer, sheet_name="Analytical Accuracy", index=False, header=False
        )
    
    IS_lims = pd.DataFrame({
        "Component": component_names,
        "Internal Standard": component_IntS_names if component_IntS_names else ["N/A"] * len(component_names),
        "Injection Standard": injection_standard_names if injection_standard_names else ["N/A"] * len(component_names),
        "Analytical Low Quant Lim [ng/L]": [lowlim[c] for c in component_names],
        "Analytical High Quant Lim [ng/L]": [highlim[c] for c in component_names],
        f"Matrix Low Quant Lim [{units}]": matrix_lowlim.min(axis=1).values,
        f"Matrix High Quant Lim [{units}]": matrix_highlim.max(axis=1).values
    })
    IS_lims.to_excel(writer, sheet_name="IS and Quant Lims", index=False)

    notes = pd.DataFrame([
        [f"All sample concentrations are reported in [{units}]"],
        ["Min LOQ set to 3x method blank or minimum calibration curve point, whichever is higher"],
        ["Continuing calibration sample recoveries are within 70%-130%, except:"],
        ["Instrument sensitivity check recoveries are within 70%-130%, except:"],
        ["Laboratory control sample recoveries are within 70%-130%, except:"],
        ["Cells highlighted yellow if surrogate recovery was outside 50-150%"],
        ["Cells highlighted red if surrogate recovery was outside 30-200%"],
        ["Additional Notes:"],
    ])
    notes.to_excel(writer, sheet_name="QA QC Notes", index=False, header=False)

    sample_compconc_tble.to_excel(writer, sheet_name="Sample Comp Conc", index=True)
    matrix_sample_conc.to_excel(writer, sheet_name="Matrix Sample Conc", index=True)

    all_sidx = pd.concat([calstd_sidx_tble, aQC_sidx_tble, mQC_sidx_tble, samples_sidx_tble], axis=1)
    sort_order = all_sidx.iloc[0].argsort()

    all_PA = pd.concat([calstd_PA_tble, aQC_PA_tble, mQC_PA_tble, samples_PA_tble], axis=1).iloc[:, sort_order]
    all_PA.to_excel(writer, sheet_name="Peak Area", index=True)

    all_IntSPA = pd.concat([calstd_IntSPA_tble, aQC_IntSPA_tble, mQC_IntSPA_tble, samples_IntSPA_tble], axis=1).iloc[:, sort_order]
    all_IntSPA.to_excel(writer, sheet_name="IS Peak Area", index=True)

    if subtract_mb and uncorrected_samples_compconc_tble is not None:
        uncorrected_samples_compconc_tble.to_excel(writer, sheet_name="Uncorrected Sample Comp Conc", index=True)

print(f"[--] Finished. Wrote results to {out_path}")