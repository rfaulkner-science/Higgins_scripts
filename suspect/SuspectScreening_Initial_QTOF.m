%% Initial Suspect Screening Data Processing for LC-QTOF-MS

% Version 2.0
% Conrad Pritchard, 8/2/2024

% This script processes raw data files from SCIEX OS (txt format) and
% identifies the suspect components of interest based on criteria for peak
% acceptance. The retention times of the suspect components of interest are
% identifued. The script then writes a txt file with updated retention
% times that can be used to create a new suspect screening processing
% method in SCIEX OS with only the components of interest and updated RTs
% with 30s integration windows.

clear
clc
tic

% Table of Content:
% 1. Set Criteria for Run Mode and Peak Acceptance (USER MODIFY)
% 2. Inputs (USER MODIFY)
% 3. Read Data Files
% 4. Extract Data
% 5. Deturmine Components of Interest
% 6. Deturmine RT of Components of Interest
% 7. Create New Suspect List Method with Components of Interest

%% 1) Set Criteria for Run Mode and Peak Acceptance
% Screen mode: only ESI -, only ESI +, or both
esimode = 'both'; %'neg' |'pos' | 'both'

% Peak selection criteria
c_me = 5; %maximum mass error (ppm)
c_w = 0.35; %maximum peak width at 50 % (min)
c_a = 2e3; %minimum peak area
c_sn = 10; %minimum signal to noise ratio
c_h = 250; %minimum peak height
c_q = 0.1; %minimum peak quality
c_b = 0.1; %maximum baseline delta / height ratio

%% 2) Inputs:
% Suspect Data File
file_name = '20240909_PRACTICE_neg_sus'; % '20240415_CP_aqneg_AFCEC_susGW'; %'20240415_CP_aqneg_AFCEC_susL'; %'20230804_CP_aqneg_AFCEC_susGW'; %'20230804_CP_aqneg_AFCEC_susL'; % '20230502_CP_negsoil_AFCEC_sus20-26'; %'20240415_CP_aqneg_AFCEC_susL'; %'20240626_CP_soilpos_AFCEC_sus1-14'; %'20240611_CP_soilpos_AFCEC_sb5_6_12'; %'20240626_CP_soilpos_AFCEC_sus15-25'; %' %'20230506_CP_soil_neg_AFCEC_sus_sb6-12'; %'PFAS NEG NIST+ suspect method_2024July_TEST'; % '20230502_CP_negsoil_AFCEC_sus20-26'; %'20230502_CP_negsoil_AFCEC_sus13-19'; %'20230506_CP_soil_neg_AFCEC_sus_sb6-12'; %'20230506_CP_soil_neg_AFCEC_sus'; %'20230424_CP_soil_neg_AFCEC_sus';
path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/AFCEC/Suspect Data/';
%path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/Forensics/';

% suspect xic list
file2_name = 'Higgins Suspect List 20240909'; %PFAS NEG NIST+ suspect method_2024July'; %; % list of all suspect compounds
path2_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/MatLab/Suspect Screening Files/';

%% 3) Read Data Files
% Read in .txt file:
rawdata_table = readtable(fullfile(path_name, [file_name '.txt']), 'Delimiter', '\t', 'VariableNamingRule', 'preserve', 'ReadVariableNames', true); % Reads in raw data from a .txt file and generates a table
rawdata_cell = table2cell(rawdata_table); % Converts the raw data table to a cell array, but does not retain column names
rawdata = [rawdata_table.Properties.VariableNames; rawdata_cell]; % Vertically concatenates column names to the cell array

% Read columns
col_samname = find(strcmp(rawdata(1, :), 'Sample Name')); %Sample name
col_samind = find(strcmp(rawdata(1, :), 'Sample Index')); %Sample index
col_injvol = find(strcmp(rawdata(1, :), 'Injection Volume')); %Injection volume
col_comname = find(strcmp(rawdata(1, :), 'Component Name')); %compound name
col_comind = find(strcmp(rawdata(1, :), 'Component Index')); %Component index
col_comgname = find(strcmp(rawdata(1, :), 'Component Group Name')); %Compound group name column
col_area = find(strcmp(rawdata(1, :), 'Area')); %Area
col_height = find(strcmp(rawdata(1, :), 'Height')); %Height
col_quality = find(strcmp(rawdata(1, :), 'Quality')); %Quality
col_RT = find(strcmp(rawdata(1, :), 'Retention Time')); %Retention time
col_width50 = find(strcmp(rawdata(1, :), 'Width at 50%')); %Width at 50 %
col_signoise = find(strcmp(rawdata(1, :), 'Signal / Noise')); %Signal to noise ratio
col_bdh = find(strcmp(rawdata(1, :), 'Baseline Delta / Height')); %Baseline delta to height ratio
col_formula = find(strcmp(rawdata(1, :), 'Formula')); %Formula
col_premass = find(strcmp(rawdata(1, :), 'Precursor Mass')); %Precursor mass
col_foundmass = find(strcmp(rawdata(1, :), 'Found At Mass')); %Found at mass
col_masserror = find(strcmp(rawdata(1, :), 'Mass Error (ppm)')); %Mass error (ppm)
col_libhit = find(strcmp(rawdata(1, :), 'Library Hit')); %Library hit
col_libscore = find(strcmp(rawdata(1, :), 'Library Score')); %Library score
col_combscore = find(strcmp(rawdata(1, :), 'Combined Score')); %Combined score
col_PAHH = find(strcmp(rawdata(1, :), 'Points Across Half Height')); %Points across half height

% Convert str to num for column with number values and NaN
wind = [col_samind col_injvol col_comind col_area col_height col_quality col_RT col_width50 col_signoise col_bdh col_premass col_foundmass col_masserror col_libscore col_combscore];

for j = 1:length(wind)
    jj = wind(j);

    for i = 2:size(rawdata, 1)

        if isa(rawdata{i, jj}, 'double')
        elseif isequal(rawdata{i, jj}, 'N/A') || isequal(rawdata{i, jj}, '< 0')
            rawdata{i, jj} = NaN;
        else
            rawdata{i, jj} = str2num(rawdata{i, jj});
        end

    end

end

% Read in suspect list .xlsx file:
XIClist_table = readtable(fullfile(path2_name, [file2_name '.xlsx']), 'VariableNamingRule', 'preserve', 'ReadVariableNames', true); % Reads in raw data from a .txt file and generates a table
XIClist_cell = table2cell(XIClist_table); % Converts the raw data table to a cell array, but does not retain column names
XIClist = [XIClist_table.Properties.VariableNames; XIClist_cell]; % Vertically concatenates column names to the cell array

%% 4) Extract Data
% Deturmine number of components
numcom = size(unique(rawdata(2:end, col_comname)), 1); %number of components
com_all = rawdata(2:numcom + 1, col_comname);

% Deturmine number of samples
samname_all = unique(rawdata(2:end, col_samname)); %list of sample names
numsam = size(samname_all, 1); %number of samples

% Create list of component group for each component
for i = 1:numcom %for each component
    comglist{i, 1} = rawdata{i + 1, col_comname}; %component name
    comglist{i, 2} = rawdata{i + 1, col_comgname}; %component group name
    comglist{i, 3} = rawdata{i + 1, col_formula}; %component formula
    comglist{i, 4} = rawdata{i + 1, col_premass}; %component mass
end

toc

%% 5) Deturmine Components of Interest
% Loop through raw data to extract list of components of interest
h = 1; %set counter for numer of final components

for i = 1:numcom %for each component
    j = 0; k = 1; %reset counters
    samrowind = find(cell2mat(rawdata(2:end, col_comind)) == i); %extract row indicies of samples of component

    while j == 0 %while j counter is set to 0 (for each sample)
        masserror = rawdata{samrowind(k) + 1, col_masserror}; %extract mass error
        width50 = rawdata{samrowind(k) + 1, col_width50}; %extract width of 50 % height
        area = rawdata{samrowind(k) + 1, col_area}; %extract area
        signoise = rawdata{samrowind(k) + 1, col_signoise}; %extract signal to noise ratio
        height = rawdata{samrowind(k) + 1, col_height}; %extract height
        quality = rawdata{samrowind(k) + 1, col_quality}; %extract quality
        bdh = rawdata{samrowind(k) + 1, col_bdh}; %extract baseline delta height ratio
        %if component has mass error > #, peak width > #min, area > #, signal to noise ratio > #, height > #cps, quality > #, and bdh < #
        if abs(masserror) < c_me && width50 < c_w && area > c_a && signoise > c_sn && height > c_h && quality > c_q && bdh < c_b
            comfinal_all{h, 1} = com_all{i};
            comfinal_all{h, 2} = i;
            h = h + 1; %add 1 to h counter
            j = 1; %exit while loop
        elseif k == numsam %if is the final sample
            j = 1; %exit while loop
        end

        k = k + 1; %add 1 to c
    end

end

toc

% Screen components of interest according to ESI Analysis Mode
%             'both'  'neg'  'pos'
%   (-) (+) |  YES            YES
%   (-)     |  YES     YES
%       (+) |                 YES
%           |  YES     YES    YES
j = 1;

for i = 1:size(comfinal_all, 1) % for each component
    XICind = find(strcmp(XIClist(:, 2), comfinal_all{i, 1})); %find XIC index

    if isempty(XICind) %if component not on Higgins Suspect List
        pos_ion_mode = {}; %set index to "" (empty)
        neg_ion_mode = {}; %set index to "" (empty)
    else
        pos_ion_mode = XIClist{XICind, 10}; % find ionization of component fom XIClist
        neg_ion_mode = XIClist{XICind, 9}; % find ionization of component fom XIClist
    end

    if isequal(esimode, 'both') && (isempty(pos_ion_mode) || ~isempty(neg_ion_mode)) % if both ESI- and does not only have esi+ symbol (i.e. either does not have a positive symbol or has a negative symbol)
        comfinal_all_new(j, :) = comfinal_all(i, :); %add component to new list
        j = j + 1;
    elseif isequal(esimode, 'neg') && isempty(pos_ion_mode) % if only ESI- and no positive symbol
        comfinal_all_new(j, :) = comfinal_all(i, :); %add component to new list
        j = j + 1;
    elseif isequal(esimode, 'pos') && (isequal(isempty(pos_ion_mode), 0) || (isempty(pos_ion_mode) && isempty(neg_ion_mode))) %if only ESI+ and has positive symbol OR no postive or negative symbol
        comfinal_all_new(j, :) = comfinal_all(i, :); %add component to new list
        j = j + 1;
    end

end

comfinal_all = comfinal_all_new; % replace list with new list

%% 6) Deturmine RT of Components of Interest
for i = 1:size(comfinal_all, 1) %for each component
    comind = comfinal_all{i, 2};
    samrowind = find(cell2mat(rawdata(2:end, col_comind)) == comind); %extract row indicies of samples of component

    for j = 1:numsam %for each sample
        masserror = rawdata{samrowind(j) + 1, col_masserror}; %extract mass error
        width50 = rawdata{samrowind(j) + 1, col_width50}; %extract width of 50 % height
        area = rawdata{samrowind(j) + 1, col_area}; %extract area
        signoise = rawdata{samrowind(j) + 1, col_signoise}; %extract signal to noise ratio
        height = rawdata{samrowind(j) + 1, col_height}; %extract height
        quality = rawdata{samrowind(j) + 1, col_quality}; %extract quality
        bdh = rawdata{samrowind(j) + 1, col_bdh}; %extract baseline delta height ratio
        %if component has mass error > #, peak width > #min, area > #, signal to noise ratio > #, height > #cps, quality > #, and bdh < #
        if abs(masserror) < c_me && width50 < c_w && area > c_a && signoise > c_sn && height > c_h && quality > c_q && bdh < c_b
            comfinal_RT_all(i, j) = rawdata{samrowind(j) + 1, col_RT}; %add peak RT to matrix
            comfinal_area_all(i, j) = log10(rawdata{samrowind(j) + 1, col_area}); %add peak area to matrix
        end

    end

    comfinal_RT_allnz = find(comfinal_RT_all(i, :) ~= 0);
    comfinal_RT(i, 1) = median(comfinal_RT_all(i, comfinal_RT_allnz));
    comfinal_area(i, 1) = max(comfinal_area_all(i, :));
end

toc

%% 7) Create New Suspect List Method with Components of Interest
% Write .txt file to be imported into SCIEX OS Method Editor
T{1, 1} = 'Non-Targeted';
T{1, 2} = 'IS';
T{1, 3} = 'Group';
T{1, 4} = 'Name';
T{1, 5} = 'Chemical Formula';
T{1, 6} = 'Isotope';
T{1, 7} = 'Adduct/Charge';
T{1, 8} = 'Gain/Loss';
T{1, 9} = 'Precursor (Q1) Mass (Da)';
T{1, 10} = 'Fragment (Q3) Mass (Da)';
T{1, 11} = 'Start - Stop';
T{1, 12} = 'XIC Width (Da)';
T{1, 13} = 'Retention Time (min)';
T{1, 14} = 'IS Name';
T{1, 15} = 'Experiment Index';
T{1, 16} = 'Min Peak Width';
T{1, 17} = 'Min Peak Height';
T{1, 18} = 'S/N Integration Threshold';
T{1, 19} = 'Smoothing Width';
T{1, 20} = 'Noise Percentage';
T{1, 21} = 'Baseline Sub. Window';
T{1, 22} = 'Peak Splitting Factor';
T{1, 23} = 'RT Window';
T{1, 24} = 'Update RT';
T{1, 25} = 'Report Largest Peak';
T{1, 26} = 'Algorithm Input Type';
T{1, 27} = 'Algorithm Name';
T{1, 28} = 'Units';
T{1, 29} = 'Use Area';
T{1, 30} = 'Regression Type';
T{1, 31} = 'Regression Weighting';
T{1, 32} = 'Signal-to-Noise Algorithm';

for i = 1:size(comfinal_all, 1) %for each component of interest
    T{i + 1, 1} = 'False';
    T{i + 1, 2} = 'False';
    T{i + 1, 3} = comglist{comfinal_all{i, 2}, 2};
    T{i + 1, 4} = comglist{comfinal_all{i, 2}, 1};
    T{i + 1, 5} = comglist{comfinal_all{i, 2}, 3};

    if isequal(esimode, 'neg') || isequal(esimode, 'both') %if esi mode is neg or both
        T{i + 1, 7} = '[M-H]-'; % Set adduct to [M-H]- Note that 'both' assumed esi - mode
    else
        T{i + 1, 6} = 1; %Set isotope to 1
        T{i + 1, 7} = '[M+H]+'; %Set adduct to [M+H]+
    end

    T{i + 1, 9} = comglist{comfinal_all{i, 2}, 4};
    T{i + 1, 10} = -1;
    T{i + 1, 12} = 0.00999999999999091;
    T{i + 1, 13} = comfinal_RT(i);
    T{i + 1, 15} = 1;
    T{i + 1, 16} = 3;
    T{i + 1, 17} = 100;
    T{i + 1, 18} = 3;
    T{i + 1, 19} = 1;
    T{i + 1, 20} = 40;
    T{i + 1, 21} = 2;
    T{i + 1, 22} = 5;
    T{i + 1, 23} = 30;
    T{i + 1, 24} = 0;
    T{i + 1, 25} = 'TRUE';
    T{i + 1, 26} = 'SingleTrace';
    T{i + 1, 27} = 'MQ4';
    T{i + 1, 29} = 'TRUE';
    T{i + 1, 30} = 'Linear';
    T{i + 1, 31} = '1 / x';
    T{i + 1, 32} = 'RelativeNoise';
end

writecell(T, fullfile(path_name, strcat(file_name, '_Method.txt')), 'Delimiter', 'tab');

%% 8) Create List of Compounds and Potential Isomers

% Create list of compounds with isomers
% Write document with list of all compounds, highlighting isomers
