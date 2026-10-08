#####
# Suspect Screening - Initial Data Processing for LC-QTOF-MS
# Rewrite of v2.0 in MATLAB, Conrad Pritchard 8/2/2024

# STANDALONE SINGLE-PASS VERSION
# TODO: add in the options for reading excel files that the matlab script does

import os
import re
import pandas as pd
import sys
import openpyxl

suspect_filename = input("Enter the name of the suspect screening file (Excel format): ")
suspect_dirname = input("Enter the name of the directory where the file is located: ")
suspect_pathname = os.path.join(suspect_dirname, suspect_filename)

xic_filename = input("Enter the name of the XIC file (Excel format): ")
xic_dirname = input("Enter the name of the directory where the file is located: ")
xic_pathname = os.path.join(xic_dirname, xic_filename)

# Define constants.
esimode = 'both' # ESI analysis mode: 'both', 'neg', or 'pos'
c_me = 5 # Mass error threshold (ppm)
c_w50 = 0.35 # Width at 50% threshold
c_area = 2000 # Area threshold
c_sn = 10 # Signal-to-noise threshold
c_h = 250 # Height threshold
c_q = 0.1 # Quality threshold
c_b = 0.1 # Baseline delta / height threshold

# Read suspect_pathname and check to make sure we have the columns we need.
req_cols = [
    'Sample Name',
    'Sample Index',
    'Injection Volume',
    'Component Name',
    'Component Index',
    'Component Group Name',
    'Area',
    'Height',
    'Quality',
    'Retention Time',
    'Width at 50%',
    'Signal / Noise',
    'Baseline Delta / Height',
    'Formula',
    'Precursor Mass',
    'Found At Mass',
    'Mass Error (ppm)',
    'Library Hit',
    'Library Score',
    'Combined Score',
    'Points Across Half Height'
]
suspect_df = pd.read_excel(suspect_pathname)
xic_df = pd.read_excel(xic_pathname)

### --- [intern] #4 in matlab script
# determine number of components and samples
s_unique_comp = suspect_df['Component Name'].unique().tolist()
s_num_unique_comp = len(s_unique_comp)
s_compdf = suspect_df.head(s_num_unique_comp)
# TODO: these two should produce the same result
# s_compdf = suspect_df[suspect_df['Sample Name'] == suspect_df.iloc[0]['Sample Name']]

s_unique_samples = suspect_df['Sample Name'].unique().tolist()
s_num_unique_samples = len(s_unique_samples)

# Create component group info
# TODO: `.values.tolist()` and `.copy()` should be the same, no?
# TODO: read docs
s_comp_grouplist = s_compdf[['Component Name', 'Component Group Name', 'Formula', 'Precursor Mass']].values.tolist()

### --- [intern] #5 in matlab script
s_compfinal = pd.DataFrame(columns=['Component Name', 'idx'])
for idx, c in s_compdf.iterrows():
    # Iterate through all samples that have this component.
    samples = suspect_df[suspect_df['Component Name'] == c['Component Name']]
    for _, s in samples.iterrows():
        mass_err = s['Mass Error (ppm)']
        w50 = s['Width at 50%']
        area = s['Area']
        signoise = s['Signal / Noise']
        height = s['Height']
        quality = s['Quality']
        bdh = s['Baseline Delta / Height']
        
        if abs(mass_err) < c_me and w50 < c_w50 and area > c_area and signoise > c_sn and height > c_h and quality > c_q and bdh < c_b:
            # Interesting way to do this.
            s_compfinal.loc[len(s_compfinal)] = [c['Component Name'], idx]

# from matlab:
# Screen components of interest according to ESI Analysis Mode
#             'both'  'neg'  'pos'
#   (-) (+) |  YES            YES
#   (-)     |  YES     YES
#       (+) |                 YES
#           |  YES     YES    YES

for idx, c in s_compfinal.iterrows():
    samples = suspect_df[suspect_df['Component Name'] == c['Component Name']]
    XIC_idx = xic-df[xic_df[2] == c['Component Name']].index
    # pim - positive ion mode
    # nim - negative ion mode
    if len(XIC_idx) == 0:
        pim = ""
        nim = ""
    else:
        pim = xic_df.iloc[XIC_idx[0], 6]
        nim = xic_df.iloc[XIC_idx[0], 5]
    
    # Skip conditions:
    if esimode == 'both' and (pim is "" or nim is not ""):
        continue
    elif esimode == 'neg' and pim is "":
        continue
    elif esimode == 'pos' and (pim is not "" or (pim is "" and nim is "")):
        continue
    else:
        # Drop.
        s_compfinal.drop(idx, inplace=True)

### --- [intern] #6 in matlab script
# Determine RT of components of interest
for idx, c in s_compfinal.iterrows():
    # Find all the samples that have this component.
    samples = suspect_df[suspect_df['Component Name'] == c['Component Name']]
    for _, s in samples.iterrows():
        pass

### --- [intern] #7 create new suspect list method with components of interest
# write as .txt file