% [intern] formatted, added self-notes
%% Final Suspect Screening Data Processing for LC-QTOF-MS

% Version 4.3
% Conrad Pritchard, 3/19/2025

% This script processes revised data files from SCIEX OS (txt format) and
% performs a semi-quantitative analsys of acceptable peaks for previously
% identified components of interest. The target data is used to develop the
% calibration curve and provide internal and surrogate standards. The
% spreadsheet CalibrantMatch is used to provide calibrants for discovered
% suspect PFASs (Note that you will need a new tab for each data set).
% This script outputs a final results file with tabs for:
%  (1) dilution-corrected, RL-adjusted, semi-quant matrix concentrations;
%  (2) semi-quant matrix concentrations;
%  (3) in-vial concentrations;
%  (4) summary of semi-quant calculation values;
%  (5) raw component peak area;
%  (6) template for assigning confidence intervals.

clear
clc
tic

% Table of Content:
% 1. Set Criteria for Peak Acceptance (USER MODIFY)
% 2. Inputs (USER MODIFY)
% 3. Read Data Files
% 4. Extract Number of Components and Samples
% 5. Deturmine Suspect Components of Interest
% 6. Extract RT, PA, etc of Suspect Components of Interest
% 7. Read in Suspect Calibrant Match File and Assign Calibrants
% 8. Extract Target Component Data
% 9. Create Target Component Calibration Curve and Range
% 10. Calculate Semi-Quant In-Vial Concentrations for Final Suspect Components
% 11. Identify Reporting Limits and Remove QCs
% 12) Calculate Matrix Conc, Matrix LoQ Range, and Replace Values Outside of Matrix LoQ
% 13) Correct Diluted Values
% 14. Seperate Target Components from Suspect Components
% 15. Construct Results File
% 16. Functions

%% 1) Set Criteria for Peak Acceptance
% Peak selection criteria - [RECOMMENDED LEAST RESTRICTIVE, MOST RESTRICTIVE]
c_me = 5; %maximum mass error (ppm) [5]
c_w = 1; %maximum peak width at 50 % (min) [0.35 1]
c_a = 2000; %minimum peak area,  [2,000 50,000]
c_sn = 10; %minimum signal to noise ratio [10]
c_h = 100; %minimum peak height [100 250]
c_q = 0.1; %minimum peak quality [0.1 0.6]
c_b = 0.1; %maximum baseline delta / height ratio [0.1 10]

%% 2) Inputs:
% [intern] Get suspect data file, target data file, calibrant match file; read suspect data columns
{
    % Suspect Data File
    sfile_name = 'enspired_solutions_POSsuspect_trimmed_may2026'; %'20230804_CP_aqneg_AFCEC_susL_Final'; %'20230804_CP_aqneg_AFCEC_susGW_Final'; %'20240626_CP_soilpos_AFCEC_sus15-25_Final'; %'20240626_CP_soilpos_AFCEC_sus1-14_Final'; %'20240611_CP_soilpos_AFCEC_sb5_6_12_Final'; %'20230506_CP_soilneg_AFCEC_sus6-12_final'; %'20230506_CP_soilneg_AFCEC_sus_final'; %'20230502_CP_negsoil_AFCEC_sus20-26_Final'; %'20230502_CP_negsoil_AFCEC_sus13-19_Final'; %'20230424_CP_soil_neg_AFCEC_sus_final'; %
    spath_name = 'C:\Users\richa\OneDrive - Colorado School of Mines\Documents\LCMS_data\Quantitation Results\'; %D:\All the stuff that was on my desktop\Not So New Folder\Graduate School Stuff\Lab Stuff\Data Processing Stuff\20230516_CP_aq_neg\
    %spath_name = '/Users/jconradpritchard/Downloads/';

    % Target Data File
    tfile_name = 'enspired_solutions_POStarget2'; %'20230804_CP_aqneg'; %'20240626_CP_soilpos'; %'20240611_CP_soilpos'; %'20230506_CP_soilneg'; %'20230502_CP_negsoil'; %'20230424_CP_soil_neg'; %
    tpath_name = 'C:\Users\richa\OneDrive - Colorado School of Mines\Documents\LCMS_data\'; %D:\All the stuff that was on my desktop\Not So New Folder\Graduate School Stuff\Lab Stuff\Data Processing Stuff\20230516_CP_aq_neg\
    %tpath_name = '/Users/jconradpritchard/Downloads/';

    % Suspect Calibrant Matching
    % This file should contain a tab with the suspect components in the data
    % set in column A, and the target calibrant in column B. No header in the
    % document. The tab should have the same name (first 31 characters - should
    % be excel default maximum name) as the suspect data set.
    file2_name = 'CalibrantMatch_fixed_ver1'; % list of calibrants for suspect compounds
    path2_name = 'C:\Users\richa\OneDrive\Documents\Research_Mines\Higgins_Group\Data_Processing\Data Processing\';
    %path2_name = '/Users/jconradpritchard/Downloads/';

    %% 3) Read Data Files
    % Read in raw suspect data .txt file:
    srawdata_table = readtable(fullfile(spath_name, [sfile_name '.txt']), 'Delimiter', '\t', 'VariableNamingRule', 'preserve', 'ReadVariableNames', true); % Reads in raw data from a .txt file and generates a table
    srawdata_cell = table2cell(srawdata_table); % Converts the raw data table to a cell array, but does not retain column names
    srawdata = [srawdata_table.Properties.VariableNames; srawdata_cell]; % Vertically concatenates column names to the cell array

    % Read columns for raw suspect data file
    scol_samname = find(strcmp(srawdata(1, :), 'Sample Name')); %Sample name
    scol_samind = find(strcmp(srawdata(1, :), 'Sample Index')); %Sample index
    scol_injvol = find(strcmp(srawdata(1, :), 'Injection Volume')); %Injection volume
    scol_comname = find(strcmp(srawdata(1, :), 'Component Name')); %compound name
    scol_comind = find(strcmp(srawdata(1, :), 'Component Index')); %Component index
    scol_comgname = find(strcmp(srawdata(1, :), 'Component Group Name')); %Compound group name column
    scol_area = find(strcmp(srawdata(1, :), 'Area')); %Area
    scol_height = find(strcmp(srawdata(1, :), 'Height')); %Height
    scol_quality = find(strcmp(srawdata(1, :), 'Quality')); %Quality
    scol_RT = find(strcmp(srawdata(1, :), 'Retention Time')); %Retention time
    scol_width50 = find(strcmp(srawdata(1, :), 'Width at 50%')); %Width at 50 %
    scol_signoise = find(strcmp(srawdata(1, :), 'Signal / Noise')); %Signal to noise ratio
    scol_bdh = find(strcmp(srawdata(1, :), 'Baseline Delta / Height')); %Baseline delta to height ratio
    scol_formula = find(strcmp(srawdata(1, :), 'Formula')); %Formula
    scol_premass = find(strcmp(srawdata(1, :), 'Precursor Mass')); %Precursor mass
    scol_foundmass = find(strcmp(srawdata(1, :), 'Found At Mass')); %Found at mass
    scol_masserror = find(strcmp(srawdata(1, :), 'Mass Error (ppm)')); %Mass error (ppm)
    scol_libhit = find(strcmp(srawdata(1, :), 'Library Hit')); %Library hit
    scol_libscore = find(strcmp(srawdata(1, :), 'Library Score')); %Library score
    scol_combscore = find(strcmp(srawdata(1, :), 'Combined Score')); %Combined score
    scol_PAHH = find(strcmp(srawdata(1, :), 'Points Across Half Height')); %Points across half height
}

% [intern] windows-only issue, fixable in python; converts string values to numbers
{
    % Convert str to num for column with number values and NaN
    swind = [scol_samind scol_injvol scol_comind scol_area scol_height scol_quality scol_RT scol_width50 scol_signoise scol_bdh scol_premass scol_foundmass scol_masserror scol_libscore scol_combscore];

    for j = 1:length(swind)
        jj = swind(j);

        for i = 2:size(srawdata, 1)

            if isa(srawdata{i, jj}, 'double')
            elseif isequal(srawdata{i, jj}, 'N/A')
                srawdata{i, jj} = NaN;
            else
                srawdata{i, jj} = str2num(srawdata{i, jj});
            end

        end

    end
}

% [intern] IRRELEVANT: stops timer from `toc`, displays it to user
toc

% [intern] read target data columns; windows-only issue of str values to num
{
    % Read in raw target data file
    trawdata_table = readtable(fullfile(tpath_name, [tfile_name '.txt']), 'Delimiter', '\t', 'VariableNamingRule', 'preserve', 'ReadVariableNames', true); % Reads in raw data from a .txt file and generates a table
    trawdata_cell = table2cell(trawdata_table); % Converts the raw data table to a cell array, but does not retain column names
    trawdata = [trawdata_table.Properties.VariableNames; trawdata_cell]; % Vertically concatenates column names to the cell array

    % Read in columns or raw target data file
    tcol_samname = find(strcmp(trawdata(1, :), 'Sample Name')); %2; %sample name column
    tcol_samind = find(strcmp(trawdata(1, :), 'Sample Index')); %2; %sample name column
    tcol_injvol = find(strcmp(trawdata(1, :), 'Injection Volume')); %19; %sample injection volume
    tcol_comname = find(strcmp(trawdata(1, :), 'Component Name')); %23; %compound name column
    tcol_comgname = find(strcmp(trawdata(1, :), 'Component Group Name')); %23; %compound group name column
    tcol_comtype = find(strcmp(trawdata(1, :), 'Component Type')); %27; %component type
    tcol_RT = find(strcmp(trawdata(1, :), 'Retention Time')); %56 compound retention time
    tcol_RTdelta = find(strcmp(trawdata(1, :), 'Retention Time Delta (min)')); %59 compound retention time
    tcol_PreMass = find(strcmp(trawdata(1, :), 'Precursor Mass')); % Precursor Mass
    tcol_MassECon = find(strcmp(trawdata(1, :), 'Mass Error Confidence')); % Precursor Mass Error Confidence
    tcol_MassError = find(strcmp(trawdata(1, :), 'Mass Error (ppm)')); % Precursor Mass Error (ppm)
    tcol_ISname = find(strcmp(trawdata(1, :), 'IS Name')); %31; %internal standard/surrogate name column
    tcol_comPA = find(strcmp(trawdata(1, :), 'Area')); %44; %compound peak area %NOTE: =40 for Windows
    tcol_ISPA = find(strcmp(trawdata(1, :), 'IS Area')); %45; %internal standard/surrogate peak area %NOTE: =54 for Windows
    tcol_actconc = find(strcmp(trawdata(1, :), 'Actual Concentration')); %37; %actual compound concentration
    tcol_ISactconc = find(strcmp(trawdata(1, :), 'IS Actual Concentration')); %38; %IS actual compound concentration
    tcol_conc = find(strcmp(trawdata(1, :), 'Calculated Concentration')); %75; %measured compound concentration %NOTE: =43 for Windows
    tcol_type = find(strcmp(trawdata(1, :), 'Sample Type')); %10; %sample type
    tcol_used = find(strcmp(trawdata(1, :), 'Used')); %Used status for cal curve

    % Convert str to num for column with number values and NaN
    twind = [tcol_injvol, tcol_RT, tcol_RTdelta, tcol_PreMass, tcol_comPA, tcol_ISPA, tcol_actconc, tcol_conc tcol_ISactconc];

    for j = 1:length(twind)
        jj = twind(j);

        for i = 2:size(trawdata, 1)

            if isa(trawdata{i, jj}, 'double')
            elseif isequal(trawdata{i, jj}, 'N/A') || isequal(trawdata{i, jj}, '< 0')
            else
                trawdata{i, jj} = str2num(trawdata{i, jj});
            end

        end

    end
}

% Deturmine if Soil or Aqueous LC Method
if trawdata{2, tcol_injvol} == 100 %if the injection volume is 100 uL, then assume soil method
    meth = 'soil';
else %otherwise assume aqueous method
    meth = 'aque';
end

toc

%% 4) Extract Number of Components and Samples
% Deturmine number of suspect components
snumcom = size(unique(srawdata(2:end, scol_comname)), 1); %number of components
scom_all = srawdata(2:snumcom + 1, scol_comname);

% Deturmine number of suspect samples
snumsam = size(unique(srawdata(2:end, scol_samname)), 1); %number of samples
ssamname_all = srawdata(linspace(2, ((snumsam - 1) * snumcom) + 2, snumsam), scol_samname);

% [intern] new same-size table with specific rows
% Create list of suspect component group for each suspect component
for i = 1:snumcom %for each suspect component
    scomglist{i, 1} = srawdata{i + 1, scol_comname}; %component name
    scomglist{i, 2} = srawdata{i + 1, scol_comgname}; %component group name
    scomglist{i, 3} = srawdata{i + 1, scol_formula}; %component formula
    scomglist{i, 4} = srawdata{i + 1, scol_premass}; %component mass
end

% Create list of target components and internal standards
i = 1; tnumcom = 0; tnumIS = 0; tcomnamelist = {};

while isequal(trawdata(1 + i, tcol_samname), trawdata(2, tcol_samname)) %while sample name is first sample name

    % [intern] if the compound type matches ..., add [compount name, internal standard used, premass] to compound-name-list
    if isequal(trawdata{i + 1, tcol_comtype}, 'Quantifiers') || isequal(trawdata{i + 1, tcol_comtype}, 'Qualifiers') %if compound is a target compound
        tnumcom = tnumcom + 1; %add one to number compounds counter
        tcomnamelist{tnumcom, 1} = trawdata{i + 1, tcol_comname}; %add compound to compound list
        tcomnamelist{tnumcom, 2} = trawdata{i + 1, tcol_ISname}; %add IS for compound to compound list
        tcomnamelist(tnumcom, 3) = trawdata(i + 1, tcol_PreMass); %add compound mass to compound list
    end

    % [intern] if this is an internal standard, add to counter
    if isequal(trawdata{i + 1, tcol_comtype}, 'Internal Standard') || isequal(trawdata{i + 1, tcol_comtype}, 'Internal Standards') %if compound is an internal standard
        tnumIS = tnumIS + 1; %add one to number IS counter
    end

    i = i + 1;
end

tnumctot = tnumcom + tnumIS; %total number of IS + compounds in target method
toc

%% 5) Deturmine Suspect Components of Interest
% Loop through raw data to extract list of components of interest
snumfcom = 0; %set counter for numer of final components
% [intern] assumed one-sample whose name has "AFFF", and should be filtered against?
sAFFFind = find(contains(ssamname_all(1, :), 'AFFF')); % find sample indicie to AFFF Bullseye

if isempty(sAFFFind)
    sAFFFind = 0;
end

for i = 1:snumcom %for each component
    j = 0; k = 1; %reset counters (j=0/1; k=sample number)
    % [intern] in all raw data, look thru samples and find indices of component (found by matching compound index) in all samples
    % [intern] assumption: every compound has its own unique index?
    samrowind = find(cell2mat(srawdata(2:end, scol_comind)) == i); %extract row indicies of samples of component

    % [intern] while true,
    while j == 0 %while j counter is set to 0 (for each sample)

        % [intern] if we've checked all sample rows with our component, break
        if k > length(samrowind)
            j = 1;
            break
        end

        % [intern] Collect specific data from each sample of our component
        % [intern] ex. sam3 of comp "PFOA" id 10; rowid 450 -- k=3 -- name "PFOA" -- id 10 -- ...details [accurate?]
        masserror = srawdata{samrowind(k) + 1, scol_masserror}; %extract mass error
        width50 = srawdata{samrowind(k) + 1, scol_width50}; %extract width of 50 % height
        area = srawdata{samrowind(k) + 1, scol_area}; %extract area
        signoise = srawdata{samrowind(k) + 1, scol_signoise}; %extract signal to noise ratio
        height = srawdata{samrowind(k) + 1, scol_height}; %extract height
        quality = srawdata{samrowind(k) + 1, scol_quality}; %extract quality
        bdh = srawdata{samrowind(k) + 1, scol_bdh}; %extract baseline delta height ratio
        %if component has mass error > #, peak width > #min, area > #, signal to noise ratio > #, height > #cps, quality > #, bdh < #, and sample is not the AFFF bullseye
        if abs(masserror) < c_me && width50 < c_w && area > c_a && signoise > c_sn && height > c_h && quality > c_q && bdh < c_b && k ~= sAFFFind
            % [intern] new final component, add the component's name (collected from the first sample), the component's index, and component group name
            snumfcom = snumfcom + 1;
            scomfinal_all{snumfcom, 1} = scom_all{i};
            scomfinal_all{snumfcom, 2} = i;
            scomfinal_all{snumfcom, 4} = scomglist{i, 2};
            j = 1; %exit while loop
        elseif k == snumsam %if is the final sample
            j = 1; %exit while loop
        end

        k = k + 1; %add 1 to c
    end

end

toc

%% 6) Extract RT, PA, etc for Suspect Components of Interest
for i = 1:size(scomfinal_all, 1)
    % [intern] Get our index, find all of its rows again [same as above]
    scomind = scomfinal_all{i, 2};
    samrowind = find(cell2mat(srawdata(2:end, scol_comind)) == scomind);

    % [intern] For each sample that exists, 
    for j = 1:snumsam
        % [intern] is this logically wrong?
        % Skip if this sample has no data for this component
        if j > length(samrowind)
            scomfinal_RT_all(i, j) = 0;
            scomfinal_area_all(i, j) = 0;
            smasserror(i, j) = 0;
            slibscore(i, j) = 0;
            continue % move to next j safely
        end

        % [intern] collect specific data from the same rows; if they match, collect final component RT data into [final component, sample num]; collect area, mass error, and lib score
        % [intern] if they don't match, default = 0
        masserror = srawdata{samrowind(j) + 1, scol_masserror};
        width50 = srawdata{samrowind(j) + 1, scol_width50};
        area = srawdata{samrowind(j) + 1, scol_area};
        signoise = srawdata{samrowind(j) + 1, scol_signoise};
        height = srawdata{samrowind(j) + 1, scol_height};
        quality = srawdata{samrowind(j) + 1, scol_quality};
        bdh = srawdata{samrowind(j) + 1, scol_bdh};

        if abs(masserror) < c_me && width50 < c_w && area > c_a && signoise > c_sn && height > c_h && quality > c_q && bdh < c_b
            scomfinal_RT_all(i, j) = srawdata{samrowind(j) + 1, scol_RT};
            scomfinal_area_all(i, j) = srawdata{samrowind(j) + 1, scol_area};
            smasserror(i, j) = masserror;
            slibscore(i, j) = srawdata{samrowind(j) + 1, scol_libscore};
        else
            scomfinal_RT_all(i, j) = 0;
            scomfinal_area_all(i, j) = 0;
            smasserror(i, j) = 0;
            slibscore(i, j) = 0;
        end

    end

    % [intern] find indices of all [final component, sample num]s that DID match if condition [not 0]
    % [intern] Find the median of all the RT values that are NOT 0
    scomfinal_RT_allnz = find(scomfinal_RT_all(i, :) ~= 0);
    scomfinal_RT(i, 1) = median(scomfinal_RT_all(i, scomfinal_RT_allnz));
    % [intern] Find highest area; repeat median strategy for mass error; find max lib score
    scomfinal_area(i, 1) = max(scomfinal_area_all(i, :));
    smasserror_allnz = find(smasserror(i, :) ~= 0);
    smasserror_all(i, 1) = median(smasserror(i, smasserror_allnz));
    slibscore_all(i, 1) = max(slibscore(i, :));
    % [intern] Find the number of components that have the same group name as this one, and write for this final component
    snumhomolog(i) = length(find(strcmp(scomfinal_all(:, 4), scomfinal_all(i, 4))));
end

toc

%% 6) Identify Potential Isomers

%% 7) Read in Suspect Calibrant Match File and Assign Calibrants
% [intern] shorten file name
if length(sfile_name) > 31
    sfile_name_short = sfile_name(1:31);
else
    sfile_name_short = sfile_name;
end

% [intern] read file, to table, idx 1 is our column names
% Read calibrant match file
CalMatch_table = readtable(fullfile(path2_name, [file2_name '.xlsx']), 'Sheet', 'NEG_ALL', 'VariableNamingRule', 'preserve', 'ReadVariableNames', true);
CalMatch_cell = table2cell(CalMatch_table); % Converts the raw data table to a cell array, but does not retain column names
CalMatch = [CalMatch_table.Properties.VariableNames; CalMatch_cell]; % Vertically concatenates column names to the cell array

% [intern] Find our component's name in the table, and save the associated calibrant (or none).
% Assign calibrants for each suspect component
for i = 1:size(scomfinal_all, 1) %for each final suspect component
    matchIdx = find(strcmp(CalMatch(:, 1), scomfinal_all{i, 1}));

    if ~isempty(matchIdx)
        scomfinal_all(i, 3) = CalMatch(matchIdx, 2);
    else
        scomfinal_all(i, 3) = {'NO_CALIBRANT'};
    end

end

toc

%% 8) Extract Target Component Data
% Categories:
%   DB: double blanks
%   cal: calibration standards
%   aQC: CCV, ISC, LB (analytical QC samples)
%   mQC: LCS, LCSD, MB (method QC samples)
%   sam: samples
% Extract actual cal conc, actual IS conc, IS peak area, measured
% concentration. If value is not a number, it is recorded as NaN. Common
% non-numerical values include 'N/A', '< 0', etc.

% Note: 'extractvalue()' is a function and is defined at the bottom of the
% script.

ncal = 0; naQC = 0; nmQC = 0; nsam = 0; %reset counters

% [intern] same deal as target script (almost; check filters), use pivots instead in pandas
for i = 2:size(trawdata, 1) %for each row in the rawdata file

    if isequal(trawdata{i, tcol_comtype}, 'Internal Standards') %if row is an internal standard (then skip)
    else
        comindex = find(strcmp(tcomnamelist(:, 1), trawdata(i, tcol_comname))); %deturmine index of compound

        if isequal(trawdata{i, tcol_type}, 'Standard') %if row is a cal standard

            if isequal(trawdata{i, tcol_comname}, tcomnamelist{1}) %if row is the first row of a sample
                ncal = ncal + 1; %add one to ncal counter (signaling new sample)
            end

            tcalISact(comindex, ncal) = extractvalue(trawdata{i, tcol_ISactconc}); %extract actual cal IS concentration
            tcalactconc(comindex, ncal) = extractvalue(trawdata{i, tcol_actconc}); %extract actual cal concentration
            tcalISPA(comindex, ncal) = extractvalue(trawdata{i, tcol_ISPA}); %extract cal IS Peak Area
            tcalconc(comindex, ncal) = extractvalue(trawdata{i, tcol_conc}); %extract computed cal concentration
            tcalPA(comindex, ncal) = extractvalue(trawdata{i, tcol_comPA}); %extract peak area
            tcalrnum(comindex, ncal) = extractvalue(trawdata{i, tcol_samind}); %extract sample index
            tcalused{comindex, ncal} = trawdata{i, tcol_used}; %extract calibration used status
        elseif contains(trawdata{i, tcol_samname}, 'CCV') || contains(trawdata{i, tcol_samname}, 'ISC') || contains(trawdata{i, tcol_samname}, 'LB') || contains(trawdata{i, tcol_samname}, 'EPA') || contains(trawdata{i, tcol_samname}, 'AFFF') || contains(trawdata{i, tcol_samname}, 'QC') %if an analytical QC

            if isequal(trawdata{i, tcol_comname}, tcomnamelist{1}) %if first sample in list
                naQC = naQC + 1; %add one to ncal counter (signaling new sample)
            end

            taQCname{1, naQC} = trawdata{i, tcol_samname}; %extract aQC name
            taQCactISconc(comindex, naQC) = extractvalue(trawdata{i, tcol_ISactconc}); %extract actual aQC IS concentration
            taQCISPA(comindex, naQC) = extractvalue(trawdata{i, tcol_ISPA}); %extract aQC IS Peak Area
            taQCconc(comindex, naQC) = extractvalue(trawdata{i, tcol_conc}); %extract aQC computed concentration
            taQCPA(comindex, naQC) = extractvalue(trawdata{i, tcol_comPA}); %extract aQC peak area
            taQCrnum(comindex, naQC) = extractvalue(trawdata{i, tcol_samind}); %extract aQC sample index
        elseif contains(trawdata{i, tcol_samname}, 'LCS') || contains(trawdata{i, tcol_samname}, 'MB') %if a method QC

            if isequal(trawdata{i, tcol_comname}, tcomnamelist{1}) %if first sample in list
                nmQC = nmQC + 1; %add one to ncal counter (signaling new sample)
            end

            tmQCname{1, nmQC} = trawdata{i, tcol_samname}; %extract mQC name
            tmQCactISconc(comindex, nmQC) = extractvalue(trawdata{i, tcol_ISactconc}); %extract actual mQC IS concentration
            tmQCISPA(comindex, nmQC) = extractvalue(trawdata{i, tcol_ISPA}); %extract mQC IS Peak Area
            tmQCconc(comindex, nmQC) = extractvalue(trawdata{i, tcol_conc}); %extract cal concentration
            tmQCPA(comindex, nmQC) = extractvalue(trawdata{i, tcol_comPA}); %extract mQC peak area
            tmQCrnum(comindex, nmQC) = extractvalue(trawdata{i, tcol_samind}); %extract aQC sample index
        elseif contains(trawdata{i, tcol_samname}, 'DB') || contains(trawdata{i, tcol_samname}, 'blank check') %if a double blank or blank check, then skip
        else %if a sample

            if isequal(trawdata{i, tcol_comname}, tcomnamelist{1}) %if first sample in list
                nsam = nsam + 1; %add one to ncal counter (signaling new sample)
            end

            tsamname{1, nsam} = trawdata{i, tcol_samname}; %extract sam name
            tsamISact(comindex, nsam) = extractvalue(trawdata{i, tcol_ISactconc}); %extract sample actual IS concentration
            tsamISPA(comindex, nsam) = extractvalue(trawdata{i, tcol_ISPA}); %extract sam IS Peak Area
            tsamconc(comindex, nsam) = extractvalue(trawdata{i, tcol_conc}); %extract sam computed concentration
            tsamPA(comindex, nsam) = extractvalue(trawdata{i, tcol_comPA}); %extract sam peak area
            tsamrnum(comindex, nsam) = extractvalue(trawdata{i, tcol_samind}); %extract sample index
        end

    end

end

toc

%% 9) Deturmine Target Component Calibration Range
% Deturmine accuracy of cal curve for each target component calibration level
[calsort, calorderindex] = sort(tcalactconc(1, :), 2); %find indicies of order (cal points are often out of order...)

for i = 1:length(calorderindex) %for each calibration level
    calorder(:, i) = tcalconc(:, calorderindex(i)); %reorder calconc matrix in order in increasing concentration % [intern] <-- this
end

% [intern] new table of (calculated concentrations / actual concentrations) of calibrations
% [intern] note calorder/calsort/e are 2D [target component index from target_first_sample, calibration_sample_num]
e = calorder ./ calsort; %find accuracy of computed cal concentrations

% Deturmine points to use for calibration curve, calibration curve, and low/high limits of calibration curve
for i = 1:size(e, 1) %for each target component
    k = 1; tcalcurve_x = []; tcalcurve_y = []; %reset counter

    for j = 1:size(e, 2) %for each calibration level

        if e(i, j) > 0.7 && e(i, j) < 1.3 && (isequal(tcalused{i, calorderindex(j)}, 'TRUE') || isequal(tcalused{i, calorderindex(j)}, 'True')) %if calibration point is used and within 30 % of the actual value
            tcalcurve_x(k) = tcalactconc(i, calorderindex(j)) / tcalISact(i, calorderindex(j)); %x values = ratio of Conc/IS Conc
            tcalcurve_y(k) = tcalPA(i, calorderindex(j)) / tcalISPA(i, calorderindex(j)); %y values = ratio of PA/IS PA
            k = k + 1; %add one to counter
        end

    end

    % [intern] Find fit of x/y, get slope, get max/min calibration level
    if sum(~isnan(tcalcurve_x)) > 0 && length(tcalcurve_x) > 2
        % calibration curve curve
        fitOptions = fitoptions('Weights', 1 ./ (tcalcurve_x .^ 2)); %use 1/x^2 weighting
        [calcurve, cal_error] = fit(tcalcurve_x(:), tcalcurve_y(:), 'poly1', fitOptions);
        tslope(i) = calcurve.p1; %calculate response factor for target components
        % min and max calibration levels
        tmincal(i) = tcalcurve_x(1) * tcalISact(i, 1); %min conc/IS conc
        tmaxcal(i) = tcalcurve_x(end) * tcalISact(i, 1); %max conc/IS conc
    end

end

toc

%% 10) Calculate Semi-Quant In-Vial Concentrations for Final Suspect Components
% Loop through all components and samples
for i = 1:size(scomfinal_all, 1) %for each final suspect component
    % [intern] get the calibration standard that they used, and find the target compound that used that calibrant
    searchname = strtrim(scomfinal_all{i, 3});
    tcalind = find(strcmp(strtrim(tcomnamelist(:, 1)), searchname)); %index of target calibrant

    if isempty(tcalind)
        warning('No Calibrant match found for component: "%s" (index %d) - skipping.', searchname, i);
        sIVconc(i, 1:size(scomfinal_area_all, 2)) = NaN; % keep row count consistent
        sresponsefactor(i) = NaN;
        slowercal(i) = NaN;
        suppercal(i) = NaN;
        continue
    end

    % [intern] The mass of the component/calibration standard; response factor = slope of targetcal; slower/supper values = min/max of targetcal
    tcommass = tcomnamelist{tcalind, 3}; %mass of target calibrant
    sresponsefactor(i) = tslope(tcalind);
    slowercal(i) = tmincal(tcalind);
    suppercal(i) = tmaxcal(tcalind);

    % [intern] Go through each sample in suspect data list
    for j = 1:size(ssamname_all, 1) %for each sample in suspect data set

        if contains(ssamname_all(j), 'AFFF') %skip quantifying AFFF bullseye
        else

            % [intern] If same suspect sample NAME is in target samples,
            if ~isempty(find(strcmp(tsamname, ssamname_all(j)))) %if is a sample in target data set
                % [intern] Set the IS-actual and IS-peakarea in targets to: [the values from the target sample that matches our suspect sample, extract the targetcal component row] REVERSED
                tISact = tsamISact(tcalind, find(strcmp(tsamname, ssamname_all(j)))); %IS actual conc in target data set
                tISPA = tsamISPA(tcalind, find(strcmp(tsamname, ssamname_all(j)))); %IS PA in target data set
            elseif ~isempty(find(strcmp(taQCname, ssamname_all(j)))) %if is a analytical QC in target data set
                % [intern] If we have same suspect sample NAME in an aQC, extract the values from [the same aQC as our suspect sample, targetcal component row] REVERSED 
                tISact = taQCactISconc(tcalind, find(strcmp(taQCname, ssamname_all(j)))); %IS actual conc in target data set
                tISPA = taQCISPA(tcalind, find(strcmp(taQCname, ssamname_all(j)))); %IS PA in target data set
            elseif ~isempty(find(strcmp(tmQCname, ssamname_all(j)))) %if is a method QC in target data set
                % [intern] same for mQC
                tISact = tmQCactISconc(tcalind, find(strcmp(tmQCname, ssamname_all(j)))); %IS actual conc in target data set
                tISPA = tmQCISPA(tcalind, find(strcmp(tmQCname, ssamname_all(j)))); %IS PA in target data set
            end

            % calculate suspect in-vial concentration: conc_sus = conc_IS * PA_sus / (PA_IS * m_targ * MW_targ/MW_sus) | (y=mx)
            % [intern] He said it's correct so it must be correct
            sIVconc(i, j) = tISact * scomfinal_area_all(i, j) / (tISPA * tslope(tcalind) * tcommass / scomglist{scomfinal_all{i, 2}, 4}); % Correct

        end

    end

end

toc

%% 11) Identify Reporting Limit and Remove QCs
% Deturmine indicies of LBs and MBs
j = 1;

for i = 1:size(ssamname_all, 1) %for each sample in suspect data set

    % [intern] make list of suspect samples that have LB/MB
    if contains(ssamname_all(i), 'LB') || contains(ssamname_all(i), 'MB') %if an LB or MB
        sQCind(j, 1) = i;
        j = j + 1;
    end

end

if j > 1 %if there is an MB or LB:
    % Calculate adjusted Reporting Limits
    for i = 1:size(scomfinal_all, 1) %for each final suspect component
        % [intern] find associated target component that used the same calstd as the final suspect component
        tcalind = find(strcmp(tcomnamelist, scomfinal_all(i, 3))); %index of target calibrant

        if isempty(tcalind)
            slowRL(i) = NaN;
            shighRL(i) = NaN;
            continue;
        end

        % [intern] calculate adjusted values
        % [intern] what the repeat?
        slowRL(i) = max(3 * max(sIVconc(i, sQCind)), tmincal(tcalind));
        shighRL(i) = tmaxcal(tcalind);
        slowRL(i) = max(3 * max(sIVconc(i, sQCind)), tmincal(find(strcmp(tcomnamelist, scomfinal_all(i, 3)))));
        shighRL(i) = tmaxcal(find(strcmp(tcomnamelist, scomfinal_all(i, 3))));
    end

    % Remove QC samples from suspect data table
    k = 1;

    for i = 1:size(ssamname_all, 1) %for each sample in the suspect data set

        if ~ismember(i, sQCind) && isequal(sum(contains(ssamname_all(i), 'AFFF')), 0)
            sIVconc2(:, k) = sIVconc(:, i);
            scomfinal_area_all2(:, k) = scomfinal_area_all(:, i);
            ssamname_all2(k) = ssamname_all(i);
            k = k + 1;
        end

    end

else %Otherwise, if there NO MB or LB: (THIS IS NOT ADVISED!!!!!!)

    for i = 1:size(scomfinal_all, 1) %for each final suspect component
        tcalind = find(strcmp(tcomnamelist, scomfinal_all(i, 3))); %index of target calibrant
        slowRL(i) = tmincal(find(strcmp(tcomnamelist, scomfinal_all(i, 3))));
        shighRL(i) = tmaxcal(find(strcmp(tcomnamelist, scomfinal_all(i, 3))));
    end

    % Remove QC samples from suspect data table
    k = 1;

    for i = 1:size(ssamname_all, 1) %for each sample in the suspect data set

        if isequal(sum(contains(ssamname_all(i), 'AFFF')), 0)
            sIVconc2(:, k) = sIVconc(:, i);
            scomfinal_area_all2(:, k) = scomfinal_area_all(:, i);
            ssamname_all2(k) = ssamname_all(i);
            k = k + 1;
        end

    end

end

%% 12) Calculate Matrix Conc, Matrix LoQ Range, and Replace Values Outside of Matrix LoQ
% [intern] similar deal to target script, slightly different
if meth == 'soil' %if soil LC method
    % Read in sample mass data
    rawsoilmass_table = readcell(fullfile(tpath_name, 'soil mass.xlsx'), 'Sheet', tfile_name);

    for k = 1:size(rawsoilmass_table, 1)

        if isnumeric(rawsoilmass_table{k, 1}) %if sample name is a number
            rawsoilmass_table{k, 1} = num2str(rawsoilmass_table{k, 1}); %convert sample name to string
        end

    end

    tmatmass = (cell2mat(rawsoilmass_table(2:end, 2)) .* (1 - cell2mat(rawsoilmass_table(2:end, 3))))'; %extract matrix mass data
    % For each suspect sample, deturmine matrix concentration and reporting limits
    for i = 1:size(ssamname_all2, 2) %for each suspect sample
        massind = find(strcmp(rawsoilmass_table(:, 1), ssamname_all2(i))) - 1; %indicie of mass spreadsheet, NOTE: tmatmass omits first row
        % Deturmine matrix extraction factor
        if isequal(rawsoilmass_table{i + 1, 4}, 'soil') %if matrix is soil
            matexfactor = 0.4/0.1 * 1.5/1000; %total vial vol (0.4mL) / extract vial vol (0.1mL) * total extract volume (1.5mL) / 1000 [mL/L] = units [L]
        elseif isequal(rawsoilmass_table{i + 1, 4}, 'dust') %otherwise if matrix is dust
            matexfactor = 0.6/0.45 * 1.5/1000 * 5/0.25; %total vial vol (0.6mL) / extract vial vol (0.45mL) * SPE extract volume (1.5mL) / 1000 [mL/L] * total extract (5mL) / SPE extract (0.25mL) = units [L]
        elseif isequal(rawsoilmass_table{i + 1, 4}, 'SPE') %otherwise if matrix is dust
            matexfactor = 0.75/0.375 * 5/1000 * 1000; %total vial vol (0.75mL) / extract vial vol (0.35mL) * SPE extract volume (5mL) / 1000 [mL/L] * 1000 [mL/L] = units [mL]
        elseif isequal(rawsoilmass_table{i + 1, 4}, 'SPEsoil') %otherwise if matrix is bambino soil
            matexfactor = 0.75/0.375 * 5/1000 * 32/2.5; %total vial vol (0.75mL) / extract vial vol (0.375mL) * SPE extract volume (5mL) / 1000 [mL/L] * total extract (32mL) / SPE extract (2.5mL) = units [L]
        end

        % Calculate matrix concentration [ng/g]
        smatsamconc(:, i) = sIVconc2(:, i) .* matexfactor ./ tmatmass(massind); %matrix concentration [ng/g] = vial conc [ng/L] * matrix extraction factor [L] / matrix mass [g]
        % Deturmine matrix upper and lower LoQ range
        smaxmatconc(:, i) = shighRL(:) .* matexfactor ./ tmatmass(massind); %upper quant limit [ng/g] = upper caliration limit [ng/L] * mat extraction factor [L] / mass extracted [g]
        sminmatconc(:, i) = slowRL(:) .* matexfactor ./ tmatmass(massind); %upper quant limit [ng/g] = reporting limit [ng/L] * mat extraction factor [L] / mass extracted [g]

        % Replace Values outside of the matrix LoQ
        for j = 1:size(scomfinal_all, 1) %for each final suspect component

            if smatsamconc(j, i) > smaxmatconc(j, i) %if matrix conc is higher than matrix LoQ
                smatconcLoQ{j, i} = sprintf('>%.2f', smaxmatconc(j, i)); %replace concentration with '> upper LoQ #'
            elseif smatsamconc(j, i) < sminmatconc(j, i) || isnan(smatsamconc(j, i)) %if matrix conc is lower than matrix LoQ
                smatconcLoQ{j, i} = sprintf('<%.2f', sminmatconc(j, i)); %replace concentration with '> lower LoQ #'
            else
                smatconcLoQ{j, i} = smatsamconc(j, i); %matrix concentraiton is within LoQ, no change necessary
            end

        end

    end

    % Define units
    units = 'ng/g';
else
    % Calculate matrix concentration [ng/L]
    smatsamconc = sIVconc2 * 1.5/0.9; % sample conc [ng/L] = vial volume (1.5mL) / volume water sample (0.9mL)

    for i = 1:size(ssamname_all2, 2) %for each sample

        %Deturmine matrix upper and lower LoQ
        smaxmatconc(:, i) = shighRL(:) * 1.5/0.9; %upper quant limit [ng/L] = upper quant limit [ng/L] * vial volume (1.5mL) / volume water sample (0.9mL)
        sminmatconc(:, i) = slowRL(:) * 1.5/0.9; %upper quant limit [ng/g] = reporting limit [ng/L] * vial volume (1.5mL) / volume water sample (0.9mL)

        % Replace Values outside of the matrix LoQ
        for j = 1:size(scomfinal_all, 1) %for each final suspect component

            if smatsamconc(j, i) > smaxmatconc(j, i) %if matrix conc is higher than matrix LoQ
                smatconcLoQ{j, i} = sprintf('>%.2f', smaxmatconc(j, i)); %replace concentration with '> upper LoQ #'
            elseif smatsamconc(j, i) < sminmatconc(j, i) || isnan(smatsamconc(j, i)) %if matrix conc is lower than matrix LoQ
                smatconcLoQ{j, i} = sprintf('<%.2f', sminmatconc(j, i)); %replace concentration with '> lower LoQ #'
            else
                smatconcLoQ{j, i} = smatsamconc(j, i); %matrix concentraiton is within LoQ, no change necessary
            end

        end

    end

    % Define units
    units = 'ng/L';
end

%% 13) Correct Diluted Values
for i = 1:size(ssamname_all2, 2) %for each sample

    if contains(ssamname_all2(1, i), '_d') && contains(ssamname_all2(1, i), 'x') %if sample is diluted (i.e. has '_d##x' in name)
        dil = extractBetween(ssamname_all2(1, i), '_d', 'x'); ndil = str2num(dil{:}); %dilution factor is the number between the '_d' and 'x'

        for j = 1:size(smatconcLoQ, 1) %for each final suspect component

            if contains(num2str(smatconcLoQ{j, i}), '>') %if the compound is above LoQ
                smatconcLoQ_dil{j, i} = sprintf('>%.2f', smaxmatconc(j, i) * ndil); %recalculate the LoQ based on the dilution factor
            elseif contains(num2str(smatconcLoQ{j, i}), '<') %if compound is below the LoQ
                smatconcLoQ_dil{j, i} = sprintf('<%.2f', sminmatconc(j, i) * ndil); %recalculate the LoQ based on dilution factor
            else %otherwise
                smatconcLoQ_dil{j, i} = cell2mat(smatconcLoQ(j, i)) * ndil; %multiple by dilution factor
            end

        end

    else %if not diluted
        smatconcLoQ_dil(:, i) = smatconcLoQ(:, i);
    end

end

%% 15) Seperate Target Compounds
j = 1; k = 1;

for i = 1:size(scomfinal_all, 1) %for each final suspect component

    if find(strcmp(tcomnamelist(:, 1), scomfinal_all(i, 1))) > 0 %if suspect component is a target component
        smatconcLoQ_dil_tar(k, :) = smatconcLoQ_dil(i, :); %final matrix target component concentration
        smatsamconc_tar(k, :) = smatsamconc(i, :); %raw matrix target component concentration
        scomfinal_tar(k, 1) = scomfinal_all(i, 1); %final target component names
        scomcalibrantfinal_tar(k, 1) = scomfinal_all(i, 3); %final target component target calibrant
        scommassfinal_tar(k, 1) = scomglist{scomfinal_all{i, 2}, 4}; %final target component mass
        scomgroupfinal_tar(k, 1) = scomglist(scomfinal_all{i, 2}, 2); %final target component mass
        scomformulafinal_tar(k, 1) = scomglist(scomfinal_all{i, 2}, 3); %final target molecular formula
        scomfinal_RT_tar(k, 1) = scomfinal_RT(i, 1); %final target RT
        smasserror_all_tar(k, 1) = smasserror_all(i, 1); %final target mass error
        snumhomolog_tar(k, 1) = snumhomolog(1, i); %final number of target homologs
        slibscore_all_tar(k, 1) = slibscore_all(i, 1); %final number of target homologs
        sIVconc2_tar(k, :) = sIVconc2(i, :); %final In-Vial target component concentration
        scomfinal_area_all2_tar(k, :) = scomfinal_area_all2(i, :); %final target component PA
        sresponsefactor_tar(k, 1) = sresponsefactor(i); %final response factor (calibration curve slope) for target components
        slowercal_tar(k, 1) = slowercal(i); %final lower calibration range for target components
        suppercal_tar(k, 1) = suppercal(i); %final upper calibration range for target components
        sminmatconc_tar(k, 1) = min(sminmatconc(i, :)); %final lower Reporting Limit for target components
        smaxmatconc_tar(k, 1) = max(smaxmatconc(i, :)); %final upper Reporting Limit for target components
        k = k + 1;
    else %otherwise if suspect component is a suspect component
        smatconcLoQ_dil_final(j, :) = smatconcLoQ_dil(i, :); %final matrix suspect component concentration
        smatsamconc_final(j, :) = smatsamconc(i, :); %raw matrix target component concentration
        scomfinal_final(j, 1) = scomfinal_all(i, 1); %final suspect component names
        scomcalibrantfinal_final(j, 1) = scomfinal_all(i, 3); %final suspect component target calibrant
        scommassfinal_final(j, 1) = scomglist{scomfinal_all{i, 2}, 4}; %final suspect component mass
        scomgroupfinal_final(j, 1) = scomglist(scomfinal_all{i, 2}, 2); %final suspect component mass
        scomformulafinal_final(j, 1) = scomglist(scomfinal_all{i, 2}, 3); %final suspect molecular formula
        scomfinal_RT_final(j, 1) = scomfinal_RT(i, 1); %final suspect RT
        smasserror_all_final(j, 1) = smasserror_all(i, 1); %final suspect mass error
        snumhomolog_final(j, 1) = snumhomolog(1, i); %final number of target homologs
        slibscore_all_final(j, 1) = slibscore_all(i, 1); %final number of target homologs
        sIVconc2_final(j, :) = sIVconc2(i, :); %final In-Vial suspect component concentration
        scomfinal_area_all2_final(j, :) = scomfinal_area_all2(i, :); %final suspect component PA
        sresponsefactor_final(j, 1) = sresponsefactor(i); %final response factor (calibration curve slope) for suspect components
        slowercal_final(j, 1) = slowercal(i); %final lower calibration range for target components
        suppercal_final(j, 1) = suppercal(i); %final upper calibration range for target components
        sminmatconc_final(j, 1) = min(sminmatconc(i, :)); %final lower Reporting Limit for target components
        smaxmatconc_final(j, 1) = max(smaxmatconc(i, :)); %final upper Reporting Limit for target components
        j = j + 1;
    end

end

%% 16) Construct Results File
% [intern] write
% Tab 1 Suspect+Target Semi-Quant Matrix Concentration Results
tab1 = {}; %define tab as cell array
tab1(1, 1) = {strcat('Suspect Compound [', units, ']')}; %define cell A1
tab1(2:size(scomfinal_final, 1) + 1, 1) = scomfinal_final(:, 1); %define compounds beginning in cell A2 extending vertically
tab1(1, 2) = {'Potential Isomer(s)'}; %define cell B1
tab1(2:end, 2) = num2cell(scommassfinal_final(:, 1)); %fill compound masses in B2 extending vertically
tab1(1, 3) = {'Formula'}; %define cell B1
tab1(2:end, 3) = scomformulafinal_final; %fill compound formula extending downwards

tab1(1, 4) = {'Precursor Mass'}; %define cell B1
tab1(2:end, 4) = num2cell(scommassfinal_final(:, 1)); %fill compound masses in B2 extending vertically
tab1(1, 5) = {'Group'}; %define cell C1
tab1(2:end, 5) = scomgroupfinal_final; %fill compound group in BC extending downwards
tab1(1, 6:size(ssamname_all2, 2) + 5) = ssamname_all2; %define samples names in cells D1 extending horizontantlly
tab1(2:end, 6:end) = smatconcLoQ_dil_final(:, :); %fill sample concentrations extending vertically (per sample) and horizontally for each sample
a = size(scomfinal_final, 1) + 3;
tab1(a, 1) = {strcat('Target Compound [', units, ']')};
tab1(a + 1:a + size(scomfinal_tar, 1), 1) = scomfinal_tar(:, 1); %define compounds beginning in cell A2 extending vertically
tab1(a, 2) = {'Potential Isomer(s)'}; %define cell B1
tab1(a + 1:end, 2) = num2cell(scommassfinal_tar(:, 1)); %fill compound masses in B2 extending vertically
tab1(a, 3) = {'Formula'}; %define cell B1
tab1(a + 1:end, 3) = scomformulafinal_tar; %fill compound formula extending downwards
tab1(a, 4) = {'Precursor Mass'}; %define cell B1
tab1(a + 1:end, 4) = num2cell(scommassfinal_tar(:, 1)); %fill compound masses in B2 extending vertically
tab1(a, 5) = {'Group'}; %define cell C1
tab1(a + 1:end, 5) = scomgroupfinal_tar; %fill compound group in BC extending downwards
tab1(a, 6:size(ssamname_all2, 2) + 5) = ssamname_all2; %define samples names in cells D1 extending horizontantlly
tab1(a + 1:end, 6:end) = smatconcLoQ_dil_tar(:, :); %fill sample concentrations extending vertically (per sample) and horizontally for each sample

writecell(tab1, fullfile(spath_name, strcat(sfile_name, '_results.xlsx')), 'Sheet', 'SemiQuant Results'); %write tab1 into results xlsx file

% Tab 2 Suspect+Target Semi-Quant Matrix Raw Results (no dilution or RL)
tab2 = {}; %define tab as cell array
tab2(1, 1) = {strcat('Suspect Compound [', units, '] (no dilution)')}; %define cell A1
tab2(2:size(scomfinal_final, 1) + 1, 1) = scomfinal_final(:, 1); %define compounds beginning in cell A2 extending vertically
tab2(1, 2) = {'Precursor Mass'}; %define cell B1
tab2(2:end, 2) = num2cell(scommassfinal_final(:, 1)); %fill compound masses in B2 extending vertically
tab2(1, 3) = {'Group'}; %define cell C1
tab2(2:end, 3) = scomgroupfinal_final; %fill compound group in BC extending downwards
tab2(1, 4:size(ssamname_all2, 2) + 3) = ssamname_all2; %define samples names in cells D1 extending horizontantlly
tab2(2:end, 4:end) = num2cell(smatsamconc_final(:, :)); %fill sample concentrations extending vertically (per sample) and horizontally for each sample
a = size(scomfinal_final, 1) + 3;
tab2(a, 1) = {strcat('Target Compound [', units, ']')};
tab2(a + 1:a + size(scomfinal_tar, 1), 1) = scomfinal_tar(:, 1); %define compounds beginning in cell A2 extending vertically
tab2(a, 2) = {'Precursor Mass'}; %define cell B1
tab2(a + 1:end, 2) = num2cell(scommassfinal_tar(:, 1)); %fill compound masses in B2 extending vertically
tab2(a, 3) = {'Group'}; %define cell C1
tab2(a + 1:end, 3) = scomgroupfinal_tar; %fill compound group in BC extending downwards
tab2(a, 4:size(ssamname_all2, 2) + 3) = ssamname_all2; %define samples names in cells D1 extending horizontantlly
tab2(a + 1:end, 4:end) = num2cell(smatsamconc_tar(:, :)); %fill sample concentrations extending vertically (per sample) and horizontally for each sample

writecell(tab2, fullfile(spath_name, strcat(sfile_name, '_results.xlsx')), 'Sheet', 'Raw Matrix Conc'); %write tab2 into results xlsx file

% Tab 3 Suspect+Target Semi-Quant In-Vial Concentration Results
tab3 = {}; %define tab as cell array
tab3(1, 1) = {'Suspect Compound [ng/L] (no dilution)'}; %define cell A1
tab3(2:size(scomfinal_final, 1) + 1, 1) = scomfinal_final(:, 1); %define compounds beginning in cell A2 extending vertically
tab3(1, 2) = {'Precursor Mass'}; %define cell B1
tab3(2:end, 2) = num2cell(scommassfinal_final(:, 1)); %fill compound masses in B2 extending vertically
tab3(1, 3) = {'Group'}; %define cell C1
tab3(2:end, 3) = scomgroupfinal_final; %fill compound group in BC extending downwards
tab3(1, 4:size(ssamname_all2, 2) + 3) = ssamname_all2; %define samples names in cells D1 extending horizontantlly
tab3(2:end, 4:end) = num2cell(sIVconc2_final(:, :)); %fill sample concentrations extending vertically (per sample) and horizontally for each sample
a = size(scomfinal_final, 1) + 3;
tab3(a, 1) = {'Target Compound [ng/L] (no dilution)'};
tab3(a + 1:a + size(scomfinal_tar, 1), 1) = scomfinal_tar(:, 1); %define compounds beginning in cell A2 extending vertically
tab3(a, 2) = {'Precursor Mass'}; %define cell B1
tab3(a + 1:end, 2) = num2cell(scommassfinal_tar(:, 1)); %fill compound masses in B2 extending vertically
tab3(a, 3) = {'Group'}; %define cell C1
tab3(a + 1:end, 3) = scomgroupfinal_tar; %fill compound group in BC extending downwards
tab3(a, 4:size(ssamname_all2, 2) + 3) = ssamname_all2; %define samples names in cells D1 extending horizontantlly
tab3(a + 1:end, 4:end) = num2cell(sIVconc2_tar(:, :)); %fill sample concentrations extending vertically (per sample) and horizontally for each sample

writecell(tab3, fullfile(spath_name, strcat(sfile_name, '_results.xlsx')), 'Sheet', 'Raw InVial Conc'); %write tab3 into results xlsx file

% Tab 4 SemiQuant Calculation
tab4 = {}; %define tab as cell array
tab4(1, 1) = {'Suspect Compound'}; %define cell A1
tab4(2:size(scomfinal_final, 1) + 1, 1) = scomfinal_final(:, 1); %define compounds beginning in cell A2 extending vertically
tab4(1, 2) = {'Precursor Mass'}; %define cell B1
tab4(2:end, 2) = num2cell(scommassfinal_final(:, 1)); %fill compound masses in B2 extending vertically
tab4(1, 3) = {'Group'}; %define cell C1
tab4(2:end, 3) = scomgroupfinal_final; %fill compound group in BC extending downwards
tab4(1, 4) = {'Calibrant'}; %define cell C1
tab4(2:end, 4) = scomcalibrantfinal_final; %fill compound calibrant in D2 extending downwards
tab4(1, 5) = {'Calibrant Response Factor'}; %define cell E1
tab4(2:end, 5) = num2cell(sresponsefactor_final); %fill calibrant response factor in E2 extending downwards
tab4(1, 6) = {'Lower Cal Range'}; %define cell F1
tab4(2:end, 6) = num2cell(slowercal_final); %fill compound lower cal range in F2 extending downwards
tab4(1, 7) = {'Upper Cal Range'}; %define cell G1
tab4(2:end, 7) = num2cell(suppercal_final); %fill compound upper cal range in G2 extending downwards
tab4(1, 8) = {'Lower RL'}; %define cell H1
tab4(2:end, 8) = num2cell(sminmatconc_final); %fill compound lower RL in H2 extending downwards
tab4(1, 9) = {'Upper RL'}; %define cell I1
tab4(2:end, 9) = num2cell(smaxmatconc_final); %fill compound upper RL in I2 extending downwards
a = size(scomfinal_final, 1) + 3;
tab4(a, 1) = {'Target Compound'};
tab4(a + 1:a + size(scomfinal_tar, 1), 1) = scomfinal_tar(:, 1); %define compounds beginning in cell A2 extending vertically
tab4(a, 2) = {'Precursor Mass'}; %define cell B1
tab4(a + 1:end, 2) = num2cell(scommassfinal_tar(:, 1)); %fill compound masses in B2 extending vertically
tab4(a, 3) = {'Group'}; %define cell C1
tab4(a + 1:end, 3) = scomgroupfinal_tar; %fill compound group in BC extending downwards
tab4(a, 4) = {'Calibrant'}; %define cell C1
tab4(a + 1:end, 4) = scomcalibrantfinal_tar; %fill compound calibrant in D2 extending downwards
tab4(a, 5) = {'Calibrant Response Factor'}; %define cell E1
tab4(a + 1:end, 5) = num2cell(sresponsefactor_tar); %fill calibrant response factor in E2 extending downwards
tab4(a, 6) = {'Lower Cal Range'}; %define cell F1
tab4(a + 1:end, 6) = num2cell(slowercal_tar); %fill compound lower cal range in F2 extending downwards
tab4(a, 7) = {'Upper Cal Range'}; %define cell G1
tab4(a + 1:end, 7) = num2cell(suppercal_tar); %fill compound upper cal range in G2 extending downwards
tab4(a, 8) = {'Lower RL'}; %define cell H1
tab4(a + 1:end, 8) = num2cell(sminmatconc_tar); %fill compound lower RL in H2 extending downwards
tab4(a, 9) = {'Upper RL'}; %define cell I1
tab4(a + 1:end, 9) = num2cell(smaxmatconc_tar); %fill compound upper RL in I2 extending downwards

writecell(tab4, fullfile(spath_name, strcat(sfile_name, '_results.xlsx')), 'Sheet', 'SemiQuant Calculation'); %write tab4 into results xlsx file

% Tab 5 Suspect+Target Raw Peak Areas
tab5 = {}; %define tab as cell array
tab5(1, 1) = {'Suspect Compound Peak Area'}; %define cell A1
tab5(2:size(scomfinal_final, 1) + 1, 1) = scomfinal_final(:, 1); %define compounds beginning in cell A2 extending vertically
tab5(1, 2) = {'Precursor Mass'}; %define cell B1
tab5(2:end, 2) = num2cell(scommassfinal_final(:, 1)); %fill compound masses in B2 extending vertically
tab5(1, 3) = {'Group'}; %define cell C1
tab5(2:end, 3) = scomgroupfinal_final; %fill compound group in BC extending downwards
tab5(1, 4:size(ssamname_all2, 2) + 3) = ssamname_all2; %define samples names in cells D1 extending horizontantlly
tab5(2:end, 4:end) = num2cell(scomfinal_area_all2_final(:, :)); %fill sample concentrations extending vertically (per sample) and horizontally for each sample
a = size(scomfinal_final, 1) + 3;
tab5(a, 1) = {'Target Compound [ng/L] (no dilution)'};
tab5(a + 1:a + size(scomfinal_tar, 1), 1) = scomfinal_tar(:, 1); %define compounds beginning in cell A2 extending vertically
tab5(a, 2) = {'Precursor Mass'}; %define cell B1
tab5(a + 1:end, 2) = num2cell(scommassfinal_tar(:, 1)); %fill compound masses in B2 extending vertically
tab5(a, 3) = {'Group'}; %define cell C1
tab5(a + 1:end, 3) = scomgroupfinal_tar; %fill compound group in BC extending downwards
tab5(a, 4:size(ssamname_all2, 2) + 3) = ssamname_all2; %define samples names in cells D1 extending horizontantlly
tab5(a + 1:end, 4:end) = num2cell(scomfinal_area_all2_tar(:, :)); %fill sample concentrations extending vertically (per sample) and horizontally for each sample

writecell(tab5, fullfile(spath_name, strcat(sfile_name, '_results.xlsx')), 'Sheet', 'Raw Peak Area'); %write tab5 into results xlsx file

% Tab 6 Confidence Level Assessment Template
tab6 = {}; %define tab as cell array
tab6(1, 1) = {'Suspect Confidence Level'}; %define cell A1
tab6(2:size(scomfinal_final, 1) + 1, 1) = scomfinal_final(:, 1); %define compounds beginning in cell A2 extending vertically
tab6(1, 2) = {'Precursor Mass'}; %define cell B1
tab6(2:end, 2) = num2cell(scommassfinal_final(:, 1)); %fill compound masses in B2 extending vertically
tab6(1, 3) = {'Group'}; %define cell C1
tab6(2:end, 3) = scomgroupfinal_final; %fill compound group in C2 extending downwards
tab6(1, 4) = {'Formula'}; %define cell D1
tab6(2:end, 4) = scomformulafinal_final; %fill compound formula extending downwards
tab6(1, 5) = {'RT'}; %define cell E1
tab6(2:end, 5) = num2cell(scomfinal_RT_final); %fill compound group in BC extending downwards
tab6(1, 6) = {'Mass Error'}; %define cell F1
tab6(2:end, 6) = num2cell(smasserror_all_final); %fill compound group in BC extending downwards
tab6(1, 7) = {'Homologes Present'}; %define cell H1
tab6(2:end, 7) = num2cell(snumhomolog_final); %fill number of compound homologs extending downwards
tab6(1, 8) = {'Library Score'}; %define cell I1
tab6(2:end, 8) = num2cell(slibscore_all_final); %fill compound group in BC extending downwards
tab6(1, 9) = {'Confidence Level'}; %define cell J1
%tab6(2:end,10)=scomcalibrantfinal_final; %fill compound group in BC extending downwards
tab6(1, 10) = {'Fragments Annotated'}; %define cell K1
%tab6(2:end,11)=scomcalibrantfinal_final; %fill compound group in BC eFxtending downwards
a = size(scomfinal_final, 1) + 3;
tab6(a, 1) = {'Target Confidence Level'}; %define cell A1
tab6(a + 1:a + size(scomfinal_tar, 1), 1) = scomfinal_tar(:, 1); %define compounds beginning in cell A2 extending vertically
tab6(a, 2) = {'Precursor Mass'}; %define cell B1
tab6(a + 1:end, 2) = num2cell(scommassfinal_tar(:, 1)); %fill compound masses in B2 extending vertically
tab6(a, 3) = {'Group'}; %define cell C1
tab6(a + 1:end, 3) = scomgroupfinal_tar; %fill compound group in C2 extending downwards
tab6(a, 4) = {'Formula'}; %define cell D1
tab6(a + 1:end, 4) = scomformulafinal_tar; %fill compound formula extending downwards
tab6(a, 5) = {'RT'}; %define cell E1
tab6(a + 1:end, 5) = num2cell(scomfinal_RT_tar); %fill compound group in BC extending downwards
tab6(a, 6) = {'Mass Error'}; %define cell F1
tab6(a + 1:end, 6) = num2cell(smasserror_all_tar); %fill compound group in BC extending downwards
tab6(a, 7) = {'Homologes Present'}; %define cell H1
tab6(a + 1:end, 7) = num2cell(snumhomolog_tar); %fill number of compound homologs extending downwards
tab6(a, 8) = {'Library Score'}; %define cell I1
tab6(a + 1:end, 8) = num2cell(slibscore_all_tar); %fill compound group in BC extending downwards
tab6(a, 9) = {'Confidence Level'}; %define cell J1
%tab6(2:end,10)=scomcalibrantfinal_tar; %fill compound group in BC extending downwards
tab6(a, 10) = {'Fragments Annotated'}; %define cell K1
%tab6(2:end,11)=scomcalibrantfinal_tar; %fill compound group in BC extending downwards
Indnum = num2cell(linspace(1, size(tab6, 1) - 1, size(tab6, 1) - 1)');
IndCol = [{'Index'}; Indnum];
tab6 = [IndCol tab6];
writecell(tab6, fullfile(spath_name, strcat(sfile_name, '_results.xlsx')), 'Sheet', 'Confidence Level Assessment'); %write tab1 into results xlsx file

%% 16) Functions
function value = extractvalue(input)

    if ischar(input)
        value = NaN;
    elseif isequal(input, [])
        value = NaN;
    else
        value = input;
    end

end
