% [intern] formatted, added self-notes

%% Data Processing for LC-QTOF-MS

% Version 5.2.1
% Conrad Pritchard, 2/12/2026

% Table of Contents:
% 1) Summary, Instructions, & Trouble shooting
% 2) Inputs & Options  ***USER MODIFICATIONS NEEDED***taskm
% 6) Identify Number of Compounds, Samples
% 7) Extract Data
% 8) Deturmine Calibration Range
% 9) Deturmine IS Peak Area Recovery
% 10) Check LCMS Analytical Accuracy
% 11) Check Method Accuracy and Deturmine Reporting limit
% 12) Calculate Matrix Conc, Matrix LoQ Range, and Replace Values Outsideof Matrix LoQ ***CHECK MATRIX CONVERSION MATH***
% 13) Correct Diluted Values
% 14) Write .xlsx Data File for Full Results
% 15) Functions

%% 1) Summary, Instructions, & Trouble Shooting:

% This script processes raw data files from SCIEX (txt or xlsx format) and
% produces an xlsx files with tabs for:
% 1) Reporting Table with Sample Matrix Concentrations
% 2) Surrogate Recovery
% 3) Method Accuracy
% 4) Analytical Accuracy
% 5) IS and Quantitation Limits
% 6) QA QC Notes
% 7) Raw Measured In-Vial Concentrations
% 8) Raw Measured Matrix Concentrations (note: no dilution corrections)
% 9) Peak Area
% 10) IS Peak Area
% 11) If correcting for MB, then non-corrected In Vial Concentrations

% INSTRUCTIONS:
% 1) Process data in Sciex and export txt file of all rows/columns. See
%    Important Notes below for necssary naming convention etc.
% 2) Update "Inputs" section of code. Note that the file type is not
%    included in the name, instead is a seperate input.
% 3) If running 100uL injection method, prepare additional spreadsheet of
%    sample masses and Injection Standards. See below for more information.
% 4) Check analytical QC sample concentrations in Section 4 ("Deturmine
%    Analysis Method and QCs (1mL vs 100uL)")
% 5) If necessary, check matrix conversion math in Section 12
% 6) Run code. Note, code may take up to ~3-20s to run.

% IMPORTANT NOTES:
% In SCIEX, you MUST:
% 1) Classify the calibration curve as "Standard".
% 2) Ensure that injection standard's component group is labeled as 'Inj
%    Stds' in the processing method.
% 3) If running diluted samples, label the samples as "_d##x" where '##'
%    indicates the dilution number, (e.g. 20 or 100 or 1000 etc)
% 4) Method QC sample naming convention: name matrix spikes samples '_MS';
%    laboratory control spikes + duplicates as 'LCS', 'LCSD', 'LLCS', or
%    'LLLCS'; method blank as 'MB'.
% 5) Analytical QC sample naming convention: name continuing calibration
%    verification samples as "CCV" or "QC", intrument sensitivity checks as
%    "ISC", laboratory blanks as "LB", EPA Bullseye must contain "EPA" or
%    "ICV", and AFFF Bulleye must contain "AFFF".
% 6) Ensure that sample names do not contain the letter sequences of the
%    names of the QCs anywhere in the name.

% IF USING 100uL INJECTION METHOD ON LCMS:
% Prepare "samplemass.xlsx" sheet so that the sample names are in the same
% order as in the SCIEX data sheet. If in doubt, check "samname" after
% running the script to see the sample order in the data sheet. The tabs
% must be the same as the datafile name. Its ok to have a set of single
% quotes around the sample names if copied from matlab 'samname'.
%
%        colA        colB            colC                colD
% row1  sample  |  mass (g)  |  moisture fraction  |  sample matrix
% row2  sam 1      mass 1       mf 1                  ('soil', 'dust', 'water')
% row3  ...

% COMMON MISTAKES:
% 1) Make sure that none of the samples have "LCS", "MB", "LB", "EPA",
%    "AFFF", "QC", "CCV", "ISC" ANYWHERE in the sample name. If so, the
%    sample will be treated as a calibration, analytical, or method qc sample.
% 2) Make sure that "samplemass.xlsx" is properly filled out and tab is
%    properly named if using 100uL injection.
% 3) Make sure samples names in "samplemass.xlsx" identically match names
%    in SCIEX exported file.
% 4) Ensure that when exporting data from SCIEX, export ALL rows and ALL
%    columns and use a '.' for decimal points.

clear %clear variables
clc %clear screen
tic %begin timer

%% 2) Inputs & Options:
% [intern] converges to "path_name/file_name.filetype" ==> path
{
    file_name = '20250916_CP_aqspe_neg'; %'20260106_SB_CDPHE_negSPE'; %'20250916_CP_aqspe_neg'; %'20251208_CP_negsoil_PFASLeach'; %'20250924_CP_soilspe_neg'; %'20250730_MM_wink_ACWPsoil'; % '20250307_CP_bs_neg_rerun'; %'20250130_CP_bs_neg'; %'20241220_CP_SPEsoil_neg'; %'20250114_EF_CP_bs_neg'; %'20240531_soilneg_Dust_new'; %'20240531_soilneg_Dust'; %'20241202_CP_SPEneg_dustonly'; %'20241202_CP_SPEneg'; %'20241030_EF_aqneg_Lysimeter'; %'20241015_CP_negSPE_CDPHE'; %'20240911_CP_aqneg'; %'20240915_CP_CDPHE_SPE_neg'; %'20230804_CP_aqneg'; %'20240626_CP_soilpos'; % '20240611_CP_soilpos'; %'20230502_CP_negsoil'; %'20240531_soilneg_Dust'; %'20240531_soilneg_Dust_copy_DK'; %'20240531_soilneg_Dust_JV'; %'20231201_CP_soilneg_AFCEC'; %'20240419_CP_aqneg'; %'20240331_CP_dust_neg'; % '20240321_CP_dust_neg'; %'20240307_CP_dust'; %'20231113_SS_CP_KR_results'; %'20231113_SS_CP_KR_results'; %231205 Raw Data_v3'; %'20231218_CP_soilneg_2'; %'20231111_CP_aqneg_results'; %'202311XX_SLJ_aqneg'; %'20231118_SS_CP_ATSDR_PrelimSugarloaf'; %'20230309_neg_soil_pfasleach'; %'20230921_CP_soilneg'; %'20230804_CP_aqneg'; %'20230426_CP_negsoil'; %'20230502_CP_negsoil'; %'20230506_CP_soilneg'; %
    %path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/PFAS LEACH/Results/'; %D:\All the stuff that was on my desktop\Not So New Folder\Graduate School Stuff\Lab Stuff\Data Processing Stuff\20230516_CP_aq_neg\
    %path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/ATSDR/Water Samples/Kelsey Prelim Water Samples/';
    %path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/ATSDR/Water Samples/';
    %path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/ATSDR/Air Sampler Calibration Study/Active Air Results/';
    %path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/ATSDR/Dust/Dust Extraction Methods/Results/Prelim experiments/';
    %path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/ATSDR/Dust/';
    %path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/AFCEC/Target Data/';
    %path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/AFCEC/Target Data/Final Desorption Isotherm/';
    %path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/CDPHE Colorado Background PFAS/Results/';
    %path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/Lysimeter/';
    %path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/Soil Treatment/Results/';
    path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/CEED Side by Side/Results/';
    %path_name = '/Users/jconradpritchard/Downloads/';

    % File type
    filetype = {'txt'}; % 'txt' or 'xlsx'
}

% Subtract MB values from Samples to Correct for Background Contamination
subMB = 0; %0=off, 1=on

%% 3) Read Data File
% Read data file
% [intern] parse
{
    if isequal(filetype, {'xlsx'}) % If an .xlsx file:
        rawdata_table = readtable(strcat(path_name, file_name, '.xlsx')); %Read in .xlsx file:
    elseif isequal(filetype, {'txt'}) % otherwise if a .txt file:
        % [intern] makes path, sets delimiter to tab, preserves column names, reads first row = column names
        rawdata_table = readtable(fullfile(path_name, [file_name '.txt']), 'Delimiter', '\t', 'VariableNamingRule', 'preserve', 'ReadVariableNames', true); % Reads in raw data from a .txt file and generates a table
    else
        fprintf('Error: File Type Incorrect')
    end
}

rawdata_cell = table2cell(rawdata_table); % Converts the raw data table to a cell array, but does not retain column names
rawdata = [rawdata_table.Properties.VariableNames; rawdata_cell]; % Vertically concatenates column names to the cell array
% [intern] effect: table -> cell

% [intern] get indexes for columns
col_samname = find(strcmp(rawdata(1, :), 'Sample Name')); %2; %sample name column
col_samind = find(strcmp(rawdata(1, :), 'Sample Index')); %2; %sample name column
col_injvol = find(strcmp(rawdata(1, :), 'Injection Volume')); %19; %sample injection volume
col_comname = find(strcmp(rawdata(1, :), 'Component Name')); %23; %compound name column
col_comgname = find(strcmp(rawdata(1, :), 'Component Group Name')); %23; %compound group name column
col_comtype = find(strcmp(rawdata(1, :), 'Component Type')); %27; %component type
col_RT = find(strcmp(rawdata(1, :), 'Retention Time')); %56 compound retention time
col_RTdelta = find(strcmp(rawdata(1, :), 'Retention Time Delta (min)')); %59 compound retention time
col_PreMass = find(strcmp(rawdata(1, :), 'Precursor Mass')); % Precursor Mass
col_MassECon = find(strcmp(rawdata(1, :), 'Mass Error Confidence')); % Precursor Mass Error Confidence
col_MassError = find(strcmp(rawdata(1, :), 'Mass Error (ppm)')); % Precursor Mass Error (ppm)
col_ISname = find(strcmp(rawdata(1, :), 'IS Name')); %31; %internal standard/surrogate name column
col_comPA = find(strcmp(rawdata(1, :), 'Area')); %44; %compound peak area %NOTE: =40 for Windows
col_ISPA = find(strcmp(rawdata(1, :), 'IS Area')); %45; %internal standard/surrogate peak area %NOTE: =54 for Windows
col_actconc = find(strcmp(rawdata(1, :), 'Actual Concentration')); %37; %actual compound concentration
col_ISactconc = find(strcmp(rawdata(1, :), 'IS Actual Concentration')); %38; %IS actual compound concentration
col_conc = find(strcmp(rawdata(1, :), 'Calculated Concentration')); %75; %measured compound concentration %NOTE: =43 for Windows
col_type = find(strcmp(rawdata(1, :), 'Sample Type')); %10; %sample type
col_used = find(strcmp(rawdata(1, :), 'Used')); %Used status for cal curve
col_polarity = find(strcmp(rawdata(1, :), 'Polarity')); %Polarity (esi mode)

%% 4) Deturmine Analysis Method and aQC Concentrations (1mL vs 100uL)
% [intern] determine method
{
    if rawdata{2, col_injvol} == 100 %if the injection volume is 100 uL and no SPE, then assume old soil or dust method
        meth = '100uL';

        if isequal(rawdata{2, col_polarity}, 'Negative')
            % Negative 100uL Analytical QC Concentrations
            CCVactconc = 250; % CCV/QC concentration [ng/L]
            ISCactconc = 25; % ISC concentration [ng/L]
            EPAactconc = 250; % EPA Bullseye [ng/L] (ICV)
        else
            % Positive 100uL Analytical QC Concentrations
            CCVactconc = 200; % CCV/QC concentration [ng/L]
            ISCactconc = 10; % ISC concentration [ng/L]
        end

    elseif rawdata{2, col_injvol} == 1000 %otherwise if large volume direct inject aqueous method
        meth = '1mL';
        % Aqueous Method QC Concentrations (1mL direct inject method)
        CCVactconc = 200; % CCV/QC concentration [ng/L]
        ISCactconc = 33.33; % ISC concentration [ng/L]
        EPAactconc = 333.33; % EPA Bullseye [ng/L]
    else
        msg = 'ERROR: unknown analysis method';
        fprintf(msg);
        return
    end
}

%% 5) If Windows, Convert str to num for Number Columns
wind = [col_injvol, col_RT, col_RTdelta, col_PreMass, col_comPA, col_ISPA, col_actconc, col_conc col_ISactconc];

for j = 1:length(wind)
    jj = wind(j);

    for i = 2:size(rawdata, 1)

        if isa(rawdata{i, jj}, 'double')
        elseif isequal(rawdata{i, jj}, 'N/A') || isequal(rawdata{i, jj}, '< 0')
        else
            rawdata{i, jj} = str2num(rawdata{i, jj});
        end

    end

end

%% 6) Identify Number of Compounds, Samples
i = 1; numcom = 0; numIS = 0; numInjS = 0; comnamelist = {}; commass = []; comgroup = {}; comISname = {}; InjSnamelist = {};

while isequal(rawdata(1 + i, col_samname), rawdata(2, col_samname)) %while sample name is first sample name

    if (isequal(rawdata{i + 1, col_comtype}, 'Quantifiers') || isequal(rawdata{i + 1, col_comtype}, 'Qualifiers')) && isequal(isequal(rawdata{i + 1, col_comgname}, 'Inj Stds'), 0) %if compound is a target compound and not an injection standard
        numcom = numcom + 1; %add one to number compounds counter
        comnamelist{numcom, 1} = rawdata{i + 1, col_comname}; %add compound to compound list
        commass(numcom, 1) = rawdata{i + 1, col_PreMass}; %add compound mass to compound list
        comgroup{numcom, 1} = rawdata{i + 1, col_comgname}; %add compound group to compound list
        comISname{numcom, 1} = rawdata{i + 1, col_ISname}; %add IS for compound to compound list
    elseif (isequal(rawdata{i + 1, col_comtype}, 'Internal Standard') || isequal(rawdata{i + 1, col_comtype}, 'Internal Standards')) && isequal(isequal(rawdata{i + 1, col_comgname}, 'Inj Stds'), 0) %if compound is an internal standard and not an injection standard
        numIS = numIS + 1; %add one to number IS counter
    elseif isequal(rawdata{i + 1, col_comgname}, 'Inj Stds') && isequal(meth, '100uL') %if an injection standard
        numInjS = numInjS + 1;
        InjSnamelist{numInjS, 1} = rawdata{i + 1, col_comname}; %add compound to compound list
    end

    i = i + 1;
end

numctot = numcom + numIS; %total number of IS + compounds in method

%% 7) Extract Data
% Categories:
%   DB: double blanks
%   cal: calibration standards
%   aQC: CCV, ISC, LB (analytical QC samples)
%   mQC: LCS, LCSD, MB (method QC samples)
%   sam: samples
% Extract actual cal conc, actual IS conc, IS peak area, measured
% concentration. If value is not a number, it is recorded as NaN. Common
% non-numerical values include 'N/A', '< 0', etc.

% Note: 'extractvalue()' is a function and is defined at the bottom of the script.

ncal = 0; naQC = 0; nmQC = 0; nsam = 0; %reset counters

for i = 2:size(rawdata, 1) %for each row in the rawdata file

    if isequal(rawdata{i, col_comtype}, 'Internal Standards') && isequal(isequal(rawdata{i, col_comgname}, 'Inj Stds'), 0) %if row is an internal standard and not an injection std then skip
    else
        comindex = find(strcmp(comnamelist(:, 1), rawdata(i, col_comname))); %deturmine index of compound

        if ~isempty(InjSnamelist)
            InjSindex = find(strcmp(InjSnamelist(:, 1), rawdata(i, col_comname))); %deturmine index of Injection Standard
        end

        if isequal(rawdata{i, col_type}, 'Standard') %if row is a cal standard

            if isequal(rawdata{i, col_comname}, comnamelist{1}) %if row is the first row of a sample
                ncal = ncal + 1; %add one to ncal counter (signaling new sample)
            end

            if isequal(rawdata{i, col_comgname}, 'Inj Stds') && isequal(meth, '100uL') %if an injection standard and 100uL method
                calInjPA(InjSindex, ncal) = extractvalue(rawdata{i, col_comPA}); %extract peak area
            elseif ~isequal(rawdata{i, col_comgname}, 'Inj Stds')
                calISact(comindex, ncal) = extractvalue(rawdata{i, col_ISactconc}); %extract actual cal IS concentration
                calactconc(comindex, ncal) = extractvalue(rawdata{i, col_actconc}); %extract actual cal concentration
                calISPA(comindex, ncal) = extractvalue(rawdata{i, col_ISPA}); %extract cal IS Peak Area
                calconc(comindex, ncal) = extractvalue(rawdata{i, col_conc}); %extract computed cal concentration
                calPA(comindex, ncal) = extractvalue(rawdata{i, col_comPA}); %extract peak area
                calrnum(comindex, ncal) = extractvalue(rawdata{i, col_samind}); %extract sample index
                calused{comindex, ncal} = rawdata{i, col_used}; %extract calibration used status
            end

        elseif contains(rawdata{i, col_samname}, 'CCV') || contains(rawdata{i, col_samname}, 'ISC') || contains(rawdata{i, col_samname}, 'LB') || contains(rawdata{i, col_samname}, 'EPA') || contains(rawdata{i, col_samname}, 'ICV') || contains(rawdata{i, col_samname}, 'QC') %if an analytical QC

            if isequal(rawdata{i, col_comname}, comnamelist{1}) %if first sample in list
                naQC = naQC + 1; %add one to ncal counter (signaling new sample)
            end

            if isequal(rawdata{i, col_comgname}, 'Inj Stds') && isequal(meth, '100uL') %if an injection standard and 100uL method
                aQCInjPA(InjSindex, naQC) = extractvalue(rawdata{i, col_comPA}); %extract peak area
            elseif ~isequal(rawdata{i, col_comgname}, 'Inj Stds')
                aQCname{1, naQC} = rawdata{i, col_samname}; %extract aQC name
                aQCactISconc(comindex, naQC) = extractvalue(rawdata{i, col_ISactconc}); %extract actual aQC IS concentration
                aQCISPA(comindex, naQC) = extractvalue(rawdata{i, col_ISPA}); %extract aQC IS Peak Area
                aQCconc(comindex, naQC) = extractvalue(rawdata{i, col_conc}); %extract aQC computed concentration
                aQCPA(comindex, naQC) = extractvalue(rawdata{i, col_comPA}); %extract aQC peak area
                aQCrnum(comindex, naQC) = extractvalue(rawdata{i, col_samind}); %extract aQC sample index
            end

        elseif contains(rawdata{i, col_samname}, 'LCS') || contains(rawdata{i, col_samname}, 'MB') %if a method QC

            if isequal(rawdata{i, col_comname}, comnamelist{1}) %if first sample in list
                nmQC = nmQC + 1; %add one to ncal counter (signaling new sample)
            end

            if isequal(rawdata{i, col_comgname}, 'Inj Stds') && isequal(meth, '100uL') %if an injection standard and 100uL method
                mQCInjPA(InjSindex, nmQC) = extractvalue(rawdata{i, col_comPA}); %extract peak area
            elseif ~isequal(rawdata{i, col_comgname}, 'Inj Stds')
                mQCname{1, nmQC} = rawdata{i, col_samname}; %extract mQC name
                mQCactISconc(comindex, nmQC) = extractvalue(rawdata{i, col_ISactconc}); %extract actual mQC IS concentration
                mQCISPA(comindex, nmQC) = extractvalue(rawdata{i, col_ISPA}); %extract mQC IS Peak Area
                mQCconc(comindex, nmQC) = extractvalue(rawdata{i, col_conc}); %extract cal concentration
                mQCPA(comindex, nmQC) = extractvalue(rawdata{i, col_comPA}); %extract mQC peak area
                mQCrnum(comindex, nmQC) = extractvalue(rawdata{i, col_samind}); %extract aQC sample index
            end

        elseif contains(rawdata{i, col_samname}, 'DB') || contains(rawdata{i, col_samname}, 'blank check') %if a double blank or blank check, then skip
        elseif isequal(contains(rawdata{i, col_samname}, 'AFFF'), 0) %if a sample (and not an AFFF bullseye)

            if isequal(rawdata{i, col_comname}, comnamelist{1}) %if first sample in list
                nsam = nsam + 1; %add one to ncal counter (signaling new sample)
            end

            if isequal(rawdata{i, col_comgname}, 'Inj Stds') && isequal(meth, '100uL') %if an injection standard and 100uL method
                samInjPA(InjSindex, nsam) = extractvalue(rawdata{i, col_comPA}); %extract peak area
            elseif ~isequal(rawdata{i, col_comgname}, 'Inj Stds')
                samname{1, nsam} = rawdata{i, col_samname}; %extract sam name
                samISact(comindex, nsam) = extractvalue(rawdata{i, col_ISactconc}); %extract sample actual IS concentration
                samISPA(comindex, nsam) = extractvalue(rawdata{i, col_ISPA}); %extract sam IS Peak Area
                samconc(comindex, nsam) = extractvalue(rawdata{i, col_conc}); %extract sam computed concentration
                samPA(comindex, nsam) = extractvalue(rawdata{i, col_comPA}); %extract sam peak area
                samrnum(comindex, nsam) = extractvalue(rawdata{i, col_samind}); %extract aQC sample index
            end

        end

    end

end

%% 8) Deturmine Calibration Range
[calsort, calorderindex] = sort(calactconc(1, :), 2); %find indicies of order (cal points are often out of order...)
%reorder calconc matrix in order in increasing concentration
for i = 1:length(calorderindex)
    calorder(:, i) = calconc(:, calorderindex(i));
end

e = calorder ./ calsort; %find accuracy of computed cal concentrations

for i = 1:size(e, 1) %for each compound
    k = 1; calcurve_x = []; calcurve_y = []; %reset counter

    for j = 1:size(e, 2) %for each cal level

        if e(i, j) > 0.7 && e(i, j) < 1.3 && (isequal(calused{i, calorderindex(j)}, 'True') || isequal(calused{i, calorderindex(j)}, 'TRUE')) %if calibration point is used and within 30 % of the actual value
            calcurve_x(k) = calsort(j); %x values = actual cal conc
            k = k + 1; %add one to counter
        end

    end

    if isempty(calcurve_x)
        lowlim(i, 1) = NaN;
        highlim(i, 1) = NaN;
    else
        lowlim(i, 1) = calcurve_x(1);
        highlim(i, 1) = calcurve_x(end);
    end

end

%% 9) Deturmine IS Peak Area Recovery
if isequal(meth, '100uL') && isequal(rawdata{2, col_polarity}, 'Negative') %if 100uL method and negative mode
    % Assign injection standard for each component
    InjSMatch = readcell(fullfile(path_name, 'samplemass.xlsx'), 'Sheet', 'NIS'); % Reads in Injection Standard Match from samplemass.xlsx document
    % Build matrix of injection standard peak areas for samples, cal, aQC, and mQC
    for i = 1:size(comnamelist, 1) %for each component
        InjSindex = find(strcmp(InjSnamelist, InjSMatch(find(strcmp(InjSMatch(:, 1), comnamelist(i))), 2))); %deturmine index of Injection Standard
        InjSlist(i, 1) = InjSnamelist(find(strcmp(InjSnamelist, InjSMatch(find(strcmp(InjSMatch(:, 1), comnamelist(i))), 2))));
        calInjMatch(i, :) = calInjPA(InjSindex, :);
        aQCInjMatch(i, :) = aQCInjPA(InjSindex, :);

        if nmQC > 0
            mQCInjMatch(i, :) = mQCInjPA(InjSindex, :);
        end

        samInjMatch(i, :) = samInjPA(InjSindex, :);
    end

    % Find IS recovery for each:
    samISrec = samISPA ./ samInjMatch ./ mean(calISPA ./ calInjMatch, 2, 'omitnan') .* (mean(calISact, 2, 'omitnan') ./ samISact); %calculate sample IS recovery, normalizing for Injection Standards and taking into account different actual IS conc values
    aQCISrec = aQCISPA ./ aQCInjMatch ./ mean(calISPA ./ calInjMatch, 2, 'omitnan') .* (mean(calISact, 2, 'omitnan') ./ aQCactISconc); %calculate aQC IS recovery, normalizing for Injection Standards and taking into account different actual IS conc values

    if nmQC > 0
        mQCISrec = mQCISPA ./ mQCInjMatch ./ mean(calISPA ./ calInjMatch, 2, 'omitnan') .* (mean(calISact, 2, 'omitnan') ./ mQCactISconc); %calculate mQC IS recovery, normalizing for Injection Standards and taking into account different actual IS conc values
    end

    % Find Injection Standard recovery for each:
    samInjRec = samInjPA ./ mean(calInjPA, 2, 'omitnan'); %calculate injection standard recovery in samples
    aQCInjRec = aQCInjPA ./ mean(calInjPA, 2, 'omitnan'); %calculate injection standard recovery in aQC

    if nmQC > 0
        mQCInjRec = mQCInjPA ./ mean(calInjPA, 2, 'omitnan'); %calculate injection standard recovery in mQC
    end

else % Otherwise, if NOT 100uL method and negative mode (i.e. 100uL pos mode, 1mL neg or pos mode)
    samISrec = samISPA ./ mean(calISPA, 2, "omitnan") .* (mean(calISact, 2) ./ samISact); %calculate sample IS recovery taking into account any difference among IS actual conc values using cal curve as reference for expected IS PA
    aQCISrec = aQCISPA ./ mean(calISPA, 2, "omitnan") .* (mean(calISact, 2) ./ aQCactISconc); %calculate sample IS recovery taking into account any difference among IS actual conc values using cal curve as reference for expected IS PA

    if nmQC > 0
        mQCISrec = mQCISPA ./ mean(calISPA, 2, "omitnan") .* (mean(calISact, 2) ./ mQCactISconc); %calculate sample IS recovery taking into account any difference among IS actual conc values using cal curve as reference for expected IS PA
    end

    %samISrec = samISPA./mean(aQCISPA,2,"omitnan").*(mean(aQCactISconc,2)./samISact); %calculate sample IS recovery taking into account any difference among IS actual conc values using QC samples as reference for expected IS PA
end

%% 10) Check LCMS Analytical Accuracy
% Calculate CCV recoveries
CCVindex = find(contains(aQCname(1, :), 'CCV')); %find the indicies of the CCVs
CCVrec = aQCconc(:, CCVindex) ./ CCVactconc;
CCVname = aQCname(:, CCVindex);

% Calculate CCV recoveries (if named QC)
QCindex = find(contains(aQCname(1, :), 'QC')); %find the indicies of the CCVs
QCrec = aQCconc(:, QCindex) ./ CCVactconc;
QCname = aQCname(:, QCindex);

% Calculate ISC recoveries
ISCindex = find(contains(aQCname(1, :), 'ISC')); %find the indicies of the ISCs
ISCrec = aQCconc(:, ISCindex) ./ ISCactconc;
ISCname = aQCname(:, ISCindex);

% Calculate LB concentrations
LBindex = find(contains(aQCname(1, :), 'LB')); %find the indicies of the LBs
LBconc = aQCconc(:, LBindex);
LBname = aQCname(:, LBindex);

% Calculate EPA Bullseye concentrations
subEPA = {'EPA', 'ICV'};
EPAindex = find(contains(aQCname(1, :), subEPA)); %find the indicies of the LBs

if ~isempty(EPAindex)
    EPArec = aQCconc(:, EPAindex) ./ EPAactconc;
    EPAname = aQCname(:, EPAindex);
else
    EPArec = [];
    EPAname = {};
end

aQCall = [CCVrec, QCrec, ISCrec, LBconc, EPArec];
aQCallname = {CCVname{:}, QCname{:}, ISCname{:}, LBname{:}, EPAname{:}};

%% 11) Check Method Accuracy and Deturmine Reporting Limit
if nmQC > 0 %if there are method QC samples
    % Calculate LCS+LCSD+LLLCSs In Vial Conc
    LCSindex = find(contains(mQCname(1, :), 'LCS')); %find the indicies of the LCS and LCSDs

    if ~isempty(LCSindex)
        LCSrec = mQCconc(:, LCSindex);
        LCSname = mQCname(:, LCSindex);
    else
        LCSrec = [];
        LCSname = {};
    end

    % Calculate MB concentrations
    MBindex = find(contains(mQCname(1, :), 'MB')); %find the indicies of the LBs
    MBconc = mQCconc(:, MBindex);
    MBname = mQCname(:, MBindex);

    % Identify Reporting Limit
    if max(MBindex) > 0 %if there is a MB sample present

        if isequal(subMB, 1) % If subtracting MB
            samconc_notMBCorr(:, :) = samconc(:, :);
        end

        for i = 1:numcom %for each compound

            if isequal(subMB, 1) % If subtracting MB
                subMBval(i, 1) = mean(MBconc(i, :), 'omitnan'); % calculate mean MB to subtract
                samconc(i, :) = samconc(i, :) - subMBval(i, 1); % subtract mean MB from Sample Conc
            else % otherwise
                subMBval(i, 1) = 0; % set mean MB to subtract as 0
            end

            if (max(MBconc(i, :)) - subMBval(i, 1)) * 3 > lowlim(i) %if the maximum modified method blank concentration is higher than the lower calibration limit
                RL(i, 1) = max(MBconc(i, :)) * 3 - subMBval(i, 1); %the reporting limit is the maximum method blank concentration x3
            else %if the maximum method blank concentration is lower than the lower calibration limit
                RL(i, 1) = lowlim(i); %otherwise the reporting limit is the lower calibration limit
            end

        end

        if max(LCSindex) > 0 %if there is a LCS/LCSD sample present
            mQCall(:, :) = [LCSrec(:, :) - subMBval(:, 1), MBconc(:, :) - subMBval(:, 1)]; %combined matrix of LCS/LCSD and MB
        else
            mQCall(:, :) = MBconc(:, :) - subMBval(:, 1); % matrix of only MB
        end

    else %if no method blank present...
        RL(i, 1) = lowlim(i); %otherwise the reporting limit is the lower calibration limit

        if max(LCSindex) > 0 %if there is a LCS/LCSD sample present
            mQCall(:, :) = LCSrec(:, :); %Matrix of only LCS/LCSD
        end

    end

else %If no method QCs...

    for i = 1:numcom %for each compound
        RL(i, 1) = lowlim(i); %otherwise the reporting limit is the lower calibration limit
    end

end

%% 12) Calculate Matrix Conc, Matrix LoQ Range, and Replace Values Outside of Matrix LoQ
if isequal(meth, '100uL') %if soil LC method
    soildata_txt = readcell(fullfile(path_name, 'samplemass.xlsx'), 'Sheet', file_name); % Reads in Injection Standard Match from samplemass.xlsx document
    soildata_num = readmatrix(fullfile(path_name, 'samplemass.xlsx'), 'Sheet', file_name); % Reads in Injection Standard Match from samplemass.xlsx document
    matmass = (soildata_num(:, 2) .* (1 - soildata_num(:, 3)))'; %extract matrix mass data

    for i = 1:nsam %for each sample

        if isempty(find(strcmp(soildata_txt(:, 1), samname(i)))) %if sample name contains an extra ' in the soildata_txt file
            soildata_txt_ind = find(strcmp(soildata_txt(:, 1), strcat("'", samname{i}, "'"))); % find soildata_txt index corresponding to sample name
        else
            soildata_txt_ind = find(strcmp(soildata_txt(:, 1), samname(i))); %find soildata_txt index corresponding to sample name
        end

        % Deturmine matrix extraction factor
        if isequal(soildata_txt{soildata_txt_ind, 4}, 'soil') %if matrix is soil (*Nickerson Soil Method*)
            matexfactor = 0.4/0.1 * 1.5/1000; %total vial vol (0.4mL) / extract vial vol (0.1mL) * total extract volume (1.5mL) / 1000 [mL/L] = units [L]
        elseif isequal(soildata_txt{soildata_txt_ind, 4}, 'dust') %otherwise if matrix is dust
            matexfactor = 0.6/0.45 * 1.5/1000 * 5/0.25; %total vial vol (0.6mL) / extract vial vol (0.45mL) * SPE extract volume (1.5mL) / 1000 [mL/L] * total extract (5mL) / SPE extract (0.25mL) = units [L]
        elseif isequal(soildata_txt{soildata_txt_ind, 4}, 'SPEwater') %otherwise if matrix is water
            matexfactor = 0.75/0.570 * 5/1000 * 1000; %total vial vol (0.75mL) / extract vial vol (0.35mL) * SPE extract volume (5mL) / 1000 [mL/L] * 1000 [mL/L] = units [L * mL/L]
        elseif isequal(soildata_txt{soildata_txt_ind, 4}, 'SPEsoil') %otherwise if matrix is bambino soil
            matexfactor = 0.75/0.570 * 5/1000 * 32/2.5; %total vial vol (0.75mL) / extract vial vol (0.375mL) * SPE extract volume (5mL) / 1000 [mL/L] * total extract (30mL) / SPE extract (2.5mL) = units [L]
        else
            error('ERROR: User defined method unknown, see samplemass.xlsx doc and lines 451-458 in the code')
        end

        % Calculate matrix concentration [ng/g]
        sammass(:, i) = samconc(:, i) * matexfactor; %total sample mass [ng] or [ng*mL/L] = vial conc [ng/L] * matrix extraction factor [L] or [mL]
        matsamconc(:, i) = sammass(:, i) ./ matmass(:, soildata_txt_ind - 1); %matrix concentration [ng/g] or [ng/L] = sample mass [ng] or [ng*mL/L] / matrix mass [g] or [mL]

        % Deturmine matrix upper and lower LoQ range
        mhighlim(:, i) = highlim * matexfactor / matmass(soildata_txt_ind - 1); %upper quant limit [ng/g] or [ng/L] = upper caliration limit [ng/L] * mat extraction factor [L] or [mL] / mass extracted [g] or [mL]
        mlowlim(:, i) = RL * matexfactor / matmass(soildata_txt_ind - 1); %upper quant limit [ng/g] = reporting limit [ng/L] * mat extraction factor [L] / mass extracted [g]

        % Replace Values outside of the matrix LoQ
        for j = 1:numcom %for each compound

            if isnan(mlowlim(j, i))
                matconcLoQ{j, i} = sprintf('No RL'); %replace concentration with 'Not Reportable'
            elseif samconc(j, i) > highlim(j) %if in-vial conc is higher than in-vial upper RL
                matconcLoQ{j, i} = sprintf('>%.2f', mhighlim(j, i)); %replace concentration with '> upper LoQ #'
            elseif samconc(j, i) < RL(j) || isnan(matsamconc(j, i)) %if in-vial conc is lower than in-vial lower RL
                matconcLoQ{j, i} = sprintf('<%.2f', mlowlim(j, i)); %replace concentration with '> lower LoQ #'
            else
                matconcLoQ{j, i} = matsamconc(j, i); %matrix concentraiton is within LoQ, no change necessary
            end

        end

    end

    % Define units
    if isequal(soildata_txt{soildata_txt_ind, 4}, 'water')
        units = 'ng/L';
    else
        units = 'ng/g';
    end

else
    % Calculate matrix concentration [ng/L]
    matsamconc = samconc * 1.5/0.87; % sample conc [ng/L] = vial volume (1.5mL) / volume water sample (0.9mL)

    for i = 1:nsam %for each sample
        %Deturmine matrix upper and lower LoQ
        mhighlim(:, i) = highlim * 1.5/0.87; %upper quant limit [ng/L] = upper quant limit [ng/L] * vial volume (1.5mL) / volume water sample (0.9mL)
        mlowlim(:, i) = RL * 1.5/0.87; %upper quant limit [ng/g] = reporting limit [ng/L] * vial volume (1.5mL) / volume water sample (0.9mL)
        % Replace Values outside of the matrix LoQ
        for j = 1:numcom %for each compound

            if matsamconc(j, i) > mhighlim(j) %if matrix conc is higher than matrix LoQ
                matconcLoQ{j, i} = sprintf('>%.2f', mhighlim(j, i)); %replace concentration with '> upper LoQ #'
            elseif matsamconc(j, i) < RL(j) || isnan(matsamconc(j, i)) %if matrix conc is lower than matrix LoQ
                matconcLoQ{j, i} = sprintf('<%.2f', mlowlim(j, i)); %replace concentration with '> lower LoQ #'
            else
                matconcLoQ{j, i} = matsamconc(j, i); %matrix concentraiton is within LoQ, no change necessary
            end

        end

    end

    % Define units
    units = 'ng/L';
end

%% 13) Correct Diluted Values
for i = 1:nsam %for each sample

    if contains(samname(i), '_d') && contains(samname(i), 'x') %if sample is diluted (i.e. has '_d##x' in name)
        dil = extractBetween(samname(i), '_d', 'x'); ndil = str2num(dil{:}); %dilution factor is the number between the '_d' and 'x'

        for j = 1:numcom %for each compound

            if contains(num2str(matconcLoQ{j, i}), '>') %if the compound is above LoQ
                dmatconcLoQ{j, i} = sprintf('>%.2f', mhighlim(j, i) * ndil); %recalculate the LoQ based on the dilution factor
            elseif contains(num2str(matconcLoQ{j, i}), '<') %if compound is below the LoQ
                dmatconcLoQ{j, i} = sprintf('<%.2f', mlowlim(j, i) * ndil); %recalculate the LoQ based on dilution factor
            else %otherwise
                dmatconcLoQ{j, i} = cell2mat(matconcLoQ(j, i)) * ndil; %multiple by dilution factor
            end

        end

    else %if not diluted
        dmatconcLoQ(:, i) = matconcLoQ(:, i);
    end

end

%% 14) Write .xlsx Data File for Full Results
% All Sample Matrix Concentrations
tab1 = {}; %define tab as cell array
tab1(1, 1) = {'Compound'}; %define cell A1
tab1(2:numcom + 1, 1) = comnamelist(:, 1); %define compounds beginning in cell A2 extending vertically
tab1(1, 2) = {'Precursor Mass'}; %define cell B1
tab1(2:end, 2) = num2cell(commass(:, 1)); %fill compound masses in B2 extending vertically
tab1(1, 3) = {'Group'}; %define cell C1
tab1(2:end, 3) = comgroup; %fill compound group in BC extending downwards
tab1(1, 4:nsam + 3) = samname; %define samples names in cells D1 extending horizontantlly
tab1(2:end, 4:end) = dmatconcLoQ(:, :); %fill sample concentrations extending vertically (per sample) and horizontally for each sample

writecell(tab1, fullfile(path_name, strcat(file_name, '_results.xlsx')), 'Sheet', 'Sample Conc'); %write tab1 into results xlsx file

% Sample IS and NIS Recovery
tab2 = {}; %define tab as cell array
tab2(1, 1) = {'Surrogate Recovery [%]'}; %define cell A1
tab2(2:numcom + 1, 1) = comnamelist(:, 1);
tab2(1, 2:nsam + 1) = samname;
tab2(2:end, 2:end) = num2cell(samISrec);

if isequal(meth, '100uL') && isequal(rawdata{2, col_polarity}, 'Negative')
    tab2(numcom + 2 + 1:numcom + 2 + size(InjSnamelist, 1), 1) = InjSnamelist;
    tab2(numcom + 2 + 1:end, 2:end) = num2cell(samInjRec);
end

writecell(tab2, fullfile(path_name, strcat(file_name, '_results.xlsx')), 'Sheet', 'Surrogate Recovery');

% Method Accuracy
tab3 = {}; %define tab as cell array
tab3(1, 1) = {'LCS/LCSD In Vial Conc [ng/L] and Method Blank Concentration [ng/L]'}; %define cell A1

if nmQC > 0 %if there are method QC samples
    tab3(2:numcom + 1, 1) = comnamelist(:, 1);
    tab3(1, 2:nmQC + 1) = [LCSname{:}, MBname];
    tab3(2:end, 2:end) = num2cell(mQCall(:, :));

    if isequal(meth, '100uL') && isequal(rawdata{2, col_polarity}, 'Negative')
        tab3(numcom + 2 + 1:numcom + 2 + size(InjSnamelist, 1), 1) = InjSnamelist;
        tab3(numcom + 2 + 1:end, 2:end) = num2cell(mQCInjRec);
    end

else
    tab3(2, 2) = {'***NO METHOD ACCURACY SAMPLES FOUND***'};
end

writecell(tab3, fullfile(path_name, strcat(file_name, '_results.xlsx')), 'Sheet', 'Method Accuracy');

% Analytical Accuracy
tab4 = {}; %define tab as cell array
tab4(1, 1) = {'Continuing Calibration Verification Accuracy [%] and Laboratory Blanks [ng/L]'}; %define cell A1

if naQC > 0 %if analytical QC samples present
    tab4(2:numcom + 1, 1) = comnamelist(:, 1);
    tab4(1, 2:size(aQCall, 2) + 1) = aQCallname;
    tab4(2:end, 2:end) = num2cell(aQCall);

    if isequal(meth, '100uL') && isequal(rawdata{2, col_polarity}, 'Negative')
        tab4(numcom + 2 + 1:numcom + 2 + size(InjSnamelist, 1), 1) = InjSnamelist;
        tab4(numcom + 2 + 1:end, 2:end) = num2cell(aQCInjRec);
    end

else
    tab4(2, 2) = {'***NO ANALYTICAL ACCURACY SAMPLES FOUND***'};
end

writecell(tab4, fullfile(path_name, strcat(file_name, '_results.xlsx')), 'Sheet', 'Analytical Accuracy');

% Compound, IS, quantitation & reporting limits
tab5 = {}; %define tab as cell array
tab5(1, 1) = {'Compound'}; %define cell A1
tab5(2:numcom + 1, 1) = comnamelist(:, 1);
tab5(1, 2) = {'Internal Standard'};
tab5(2:end, 2) = comISname(:, 1);
tab5(1, 3) = {'Injection Standard'};

if ~isempty(InjSnamelist)
    tab5(2:end, 3) = InjSlist(:, 1);
end

tab5(1, 4:8) = {'Analytical Low Quant Lim [ng/L]', 'Analytical High Quant Lim [ng/L]', 'Method Low Quant Limit [ng/L]', strcat('Matrix Low Quant Lim [', units, ']'), strcat('Matrix High Qant Lim [', units, ']')};
tab5(2:end, 4:8) = num2cell([lowlim highlim RL min(mlowlim, [], 2) max(mhighlim, [], 2)]);

writecell(tab5, fullfile(path_name, strcat(file_name, '_results.xlsx')), 'Sheet', 'IS and Quant Lims');

% QC Notes for Final Data File
tab6 = {}; %define tab as cell array
tab6(1, 1) = {'QA QC Notes'}; %define cell A1
tab6(2, 1) = {strcat('All sample concentrations are reported in [', units, ']')};
tab6(3, 1) = {'Min LOQ set to 3x method blank or minimum calibration curve point, whichever is higher'};
tab6(4, 1) = {'Continuing calibration sample recoveries are within 70%-130%, except:'};
tab6(5, 1) = {'Instrument sensitivity check recoveries are within 70%-130%, except:'};
tab6(6, 1) = {'Laboratory control sample recoveries are within 70%-130%, except:'};
tab6(7, 1) = {'Cells highlighted yellow if surrorgate recovery was outside 50-150%'};
tab6(8, 1) = {'Cells highlighted red if surrorgate recovery was outside 30-200%'};
tab6(9, 1) = {'Additional Notes:'};

writecell(tab6, fullfile(path_name, strcat(file_name, '_results.xlsx')), 'Sheet', 'QA QC Notes');

% In vial Concentrations [ng/L]
tab7 = {}; %define tab as cell array
tab7(1, 1) = {'Measured In Vial Concentration [ng/L]'}; %define cell A1
tab7(2:numcom + 1, 1) = comnamelist(:, 1);
tab7(1, 2:size(samname, 2) + 1) = samname;
tab7(2:end, 2:end) = num2cell(samconc);

writecell(tab7, fullfile(path_name, strcat(file_name, '_results.xlsx')), 'Sheet', 'Raw Vial Conc');

% Raw Matrix Concentrations [ng/L] of [ng/g]
tab8 = {}; %define tab as cell array
tab8(1, 1) = {strcat('Measured Matrix Concentration [', units, '] (NO DILUTION CORRECTIONS)')}; %define cell A1
tab8(2:numcom + 1, 1) = comnamelist(:, 1);
tab8(1, 2:size(samname, 2) + 1) = samname;
tab8(2:end, 2:end) = num2cell(matsamconc);

writecell(tab8, fullfile(path_name, strcat(file_name, '_results.xlsx')), 'Sheet', 'Raw Matrix Conc');

% Raw Peak Areas, IS Peak Areas, and Injection Standard Peak Areas of all cal standards, QCs, and samples
calname = num2cell(calactconc(1, :)); %make cell array of cal names

if nmQC > 0
else
    mQCrnum = [];
    mQCname = [];
    mQCPA = [];
    mQCISPA = [];
    mQCInjPA = [];
end

allrnum = [calrnum, aQCrnum, mQCrnum, samrnum];
allname = [calname, aQCname, mQCname, samname];
allPA = [calPA, aQCPA, mQCPA, samPA];
allISPA = [calISPA, aQCISPA, mQCISPA, samISPA];

if ~isempty(InjSnamelist)
    allInjPA = [calInjPA, aQCInjPA, mQCInjPA, samInjPA];
end

[allrnum_sorted, allrnum_index] = sort(allrnum, 2);

for i = 1:size(allrnum, 2)
    allname_sorted(:, i) = allname(:, allrnum_index(1, i));
    allPA_sorted(:, i) = allPA(:, allrnum_index(1, i));
    allISPA_sorted(:, i) = allISPA(:, allrnum_index(1, i));

    if ~isempty(InjSnamelist)
        allInjPA_sorted(:, i) = allInjPA(:, allrnum_index(1, i));
    end

end

% Make PA tab
tab9 = {}; %define tab as cell array
tab9(1, 1) = {'Raw Peak Areas'}; %define cell A1
tab9(2:numcom + 1, 1) = comnamelist(:, 1);
tab9(1, 2:size(allname, 2) + 1) = allname_sorted;
tab9(2:end, 2:end) = num2cell(allPA_sorted);

if isequal(meth, '100uL') && isequal(rawdata{2, col_polarity}, 'Negative')
    tab9(numcom + 2 + 1:numcom + 2 + size(InjSnamelist, 1), 1) = InjSnamelist;
    tab9(numcom + 2 + 1:end, 2:end) = num2cell(allInjPA_sorted);
end

writecell(tab9, fullfile(path_name, strcat(file_name, '_results.xlsx')), 'Sheet', 'Peak Area');

% Make ISPA tab
tab10 = {}; %define tab as cell array
tab10(1, 1) = {'Raw Peak Areas'}; %define cell A1
tab10(2:numcom + 1, 1) = comnamelist(:, 1);
tab10(1, 2:size(allname, 2) + 1) = allname_sorted;
tab10(2:end, 2:end) = num2cell(allISPA_sorted);

if isequal(meth, '100uL') && isequal(rawdata{2, col_polarity}, 'Negative')
    tab10(numcom + 2 + 1:numcom + 2 + size(InjSnamelist, 1), 1) = InjSnamelist;
    tab10(numcom + 2 + 1:end, 2:end) = num2cell(allInjPA_sorted);
end

writecell(tab10, fullfile(path_name, strcat(file_name, '_results.xlsx')), 'Sheet', 'IS Peak Area');

if isequal(subMB, 1) %If correction for Method Blank, then include additional tab without MB Correction
    % In vial Concentrations [ng/L]
    tab11 = {}; %define tab as cell array
    tab11(1, 1) = {'Measured In Vial Concentration [ng/L] *NOT METHOD BLANK CORRECTED*'}; %define cell A1
    tab11(2:numcom + 1, 1) = comnamelist(:, 1);
    tab11(1, 2:size(samname, 2) + 1) = samname;
    tab11(2:end, 2:end) = num2cell(samconc_notMBCorr);

    writecell(tab11, fullfile(path_name, strcat(file_name, '_results.xlsx')), 'Sheet', 'Original Raw Vial Conc');
end

toc

%% 15) Functions
function value = extractvalue(input)

    if ischar(input)
        value = NaN;
    elseif isequal(input, [])
        value = NaN;
    else
        value = input;
    end

end
