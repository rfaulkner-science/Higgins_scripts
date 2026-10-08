%% Data Processing for LC-OrbiTrap-MS

% Version 2.0.1
% Conrad Pritchard, 3/4/2024

% Table of Contents:
% 1) Summary, Instructions, & Trouble shooting
% 2) Inputs **USER MODIFICATIONS NEEDED**
% 3) Set QC Sample Concentrations
% 4) Read Data File
% 5) Convert strings to numbers
% 6) Identify number of samples, sample list
% 7) Extract Data
% 8) Deturmine Calibration Range
% 9) Deturmine IS Peak Area Recovery
% 10) Check LCMS Analytical Accuracy
% 11) Check Method Accuracy and Deturmine Reporting Limit
% 12) Calculate Matrix Conc, Matrix LoQ Range, and Replace Values Outside of Matrix LoQ
% 13) Correct Diluted Values
% 14) Write .xlsx Data File for Full Results
% 15) Functions

%% 1) Summary, Instructions, & Trouble Shooting:
% This script processes raw data files from TraceFinder (xlsx format) and
% produces an xlsx files with tabs for:
% 1) Reporting Table with Sample Matrix Concentrations
% 2) Surrogate Recovery
% 3) Method Accuracy
% 4) Analytical Accuracy
% 5) Quantitation Limits
% 6) QC Information
% 7) QA QC Notes
% 8) Raw Measured In-Vial Concentrations
% 9) Raw Measured Matrix Concentrations (note: no dilution corrections)

% INSTRUCTIONS:
% 1) Process data in TraceFinder and export xlsx file of all rows/columns. See
%    Important Notes below for necssary naming convention etc.
% 2) Update "Inputs" section of code. Note that the file type is not
%    included in the name, instead is a seperate input.
% 3) If running soil or dust method, prepare additional spreadsheet of sample
%    masses. See below for more information.
% 4) Check QC sample concentrations in Section 4 ("Deturmine if soil or
%    aqueous LC method")
% 5) Run code. Note, code may take up to ~20s to run.

% IMPORTANT NOTES:
% In TraceFinder, you MUST:
% 1) Classify the calibration curve as "Standard".
% 2) Label M2PFOA as "Qualifiers" or "Quantifiers", NOT as an internal
%    standard. Note that in some data proceessing
%    methods in TraceFinder, M2PFOA is labeled as an "Internal Standard". Uncheck
%    that box inorder to ensure M2PFOA is labeled correctly.
% 3) If running diluted samples, label the samples as "_d##x" where '##'
%    indicates the dilution number, (e.g. 20 or 100 or 1000 etc)
% 4) Method QC sample naming convention: name matrix spikes samples "_MS";
%    laboratory control spikes + duplicates as LCS and LCSD.
% 5) Analytical QC sample naming convention: name continuing calibration
%    verification samples as "CCV" or "QC", intrument sensitivity checks as
%    "ISC", laboratory blanks as "LB", EPA Bullseye must contain "EPA", and
%    AFFF Bulleye must contain "AFFF".

% IF RUNNING SOIL or DUST:
% Prepare "soil mass.xlsx" sheet so that the sample names are in the same
% order as in the TraceFinder data sheet. If in doubt, check "samname" after
% running the script to see the sample order in the data sheet. The tabs
% must be the same as the datafile name.
%
%        colA        colB            colC
% row1  sample  |  mass (g)  |  moisture fraction
% row2  sam 1      mass 1       mf 1
% row3  ...

% COMMON MISTAKES:
% 1) Make sure that none of the samples have "LCS", "MB", "LB", "cal", "EPA",
%    "AFFF", "QC", "CCV", "ISC" ANYWHERE in the sample name. If so, the
%    sample will be treated as a calibration, analytical, or method qc sample.
% 2) Make sure that all sample types are properly labeled. LCS/LCS Dup and
%    MS are commonly mislabeled and need be labeled as sample type "Uknown".
% 3) Make sure that Soil Mass sheet is properly filled out and tab is
%    properly labeled.
% 4) Make sure M2PFOA is not an internal standard. Some Sciex OS quantitation
%    methods will include M2PFOA as an internal standard, so you must use afree
%    different quant method or classify it as not an internal standard in
%    the method
% 5) Make sure samples are in correct order in "soil mass.xlsx" sheet to
%    match the order of "samname"
% 6) Ensure that when exporting data from SCIEX, export ALL rows and ALL
%    columns and use a '.' for decimal points.

clear %clear variables
clc %clear screen
tic %begin timer

%% 2) Inputs:
%file_name = '20231023_CP_Run7_target_FINAL';
%file_name = '20231023_CP_Run6_target_FINAL';
%file_name = '20231023_CP_Run5_target_FINAL';
%file_name = '20231023_CP_Run4_target_FINAL';
%file_name = '20231005_CP_Run1_target_FINAL';
%file_name = '20230828_CP_Plasma_Test_Batch_FINAL';
%path_name = '/Users/jconradpritchard/Documents/Colorado School of Mines/Projects/Plasma/Results/Target/';
file_name = 'CP_UltraShortChain_20251028';
path_name = 'C:\Users\delka\Dropbox\Grad School\';
NIS_file = 'NISmatch';
NIS_path = 'C:\Users\delka\Dropbox\Higgins Research Group Shared Files\Data Processing\Orbitrap Target & Suspect SOP - Sarah\';
full_file_path = strcat(path_name, file_name, '.xlsx');
disp(full_file_path);
% Matrix/method
matmethod = 'aqueous'; %dust, soil, aqueous, plasma

%% 3) Set QC Sample Concentrations
% % Surrogate, ISC, and QC Limits
% if isequal(matmethod,'plasma') || isequal(matmethod,'soil') || isequal(matmethod,'dust')
%     CCVactconc = 200; ISCactconc = 25; LCSactconc = 750; ICVactconc = 333;
% elseif isequal(matmethod,'aqueous')
%     CCVactconc = 200; ISC1actconc = 0.07; ISC2actconc = 0.33; ISC3actconc = 0.67 ; ISC4actconc = 1.33; ISC5actconc = 1.19 ; LCSactconc = 200; ICVactconc = 333.33;
% end

%% 4) Read Data File
% Read data file
%[~,~,rawdata] = xlsread(strcat(path_name,file_name,'.xlsx'),'All Samples'); %,file_name); %read in LC TrOCs data
rawdata_table = readtable(strcat(path_name, file_name, '.xlsx'), 'VariableNamingRule', 'preserve');
rawdata_cell = table2cell(rawdata_table); % Converts the raw data table to a cell array, but does not retain column names
rawdata = [rawdata_table.Properties.VariableNames; rawdata_cell]; % Vertically concatenates column names to the cell array

col_samname = find(strcmp(rawdata(1, :), 'Filename')); %sample name column
col_samtype = find(strcmp(rawdata(1, :), 'Sample Type')); %sample type
col_comname = find(strcmp(rawdata(1, :), 'Compound')); %compound name column
col_comtype = find(strcmp(rawdata(1, :), 'Type')); %compound type (Target Compound or Internal Standard)
col_RT = find(strcmp(rawdata(1, :), 'RT')); %compound retention time
col_actRT = find(strcmp(rawdata(1, :), 'Actual RT')); %actual compound retention time
col_RTdelta = find(strcmp(rawdata(1, :), 'RT Delta')); %compound retention time delta
col_comPA = find(strcmp(rawdata(1, :), 'Area')); %44; %compound peak area %NOTE: =40 for Windows
col_conc = find(strcmp(rawdata(1, :), 'Calculated Amt')); %measured compound concentration
col_actconc = find(strcmp(rawdata(1, :), 'Theoretical Amt')); %actual compound concentration
col_mzexp = find(strcmp(rawdata(1, :), 'm/z (Expected)')); %expected m/z
col_mzapex = find(strcmp(rawdata(1, :), 'm/z (Apex)')); %apex m/z
col_mzdelta = find(strcmp(rawdata(1, :), 'm/z (Delta)')); %m/z delta
col_ISPA = find(strcmp(rawdata(1, :), 'ISTD Response')); %internal standard/surrogate peak area
col_ISact = find(strcmp(rawdata(1, :), 'ISTD Amt')); %actual internal standard/surrogate concentration
col_units = find(strcmp(rawdata(1, :), 'Final Units')); %find unitscolum
units = rawdata{2, col_units}; %find unit

rawdata_table2 = readtable(strcat(NIS_path, NIS_file, '.xlsx'), 'VariableNamingRule', 'preserve');
rawdata_cell2 = table2cell(rawdata_table2); % Converts the raw data table to a cell array, but does not retain column names
rawdata2 = [rawdata_table2.Properties.VariableNames; rawdata_cell2]; % Vertically concatenates column names to the cell array

col_IS2 = find(strcmp(rawdata2(1, :), 'Internal Standard')); %EIS column
col_NIS = find(strcmp(rawdata2(1, :), 'NIS Standard')); %NIS compound column
col_com = find(strcmp(rawdata2(1, :), 'Compound'));

%% 5) Convert str to num for Number Columns
wind = [col_RT, col_actRT, col_RTdelta, col_comPA, col_conc, col_actconc, col_mzexp, col_mzapex, col_ISPA, col_ISact];

for j = 1:length(wind)
    jj = wind(j);

    for i = 2:size(rawdata, 1)

        if isa(rawdata{i, jj}, 'double')
        elseif isequal(rawdata{i, jj}, 'N/A') || isequal(rawdata{i, jj}, '< 0') || isequal(rawdata{i, jj}, 'N/F')
        else
            rawdata{i, jj} = str2num(rawdata{i, jj});
        end

    end

end

%% 6) Identify number of samples, sample list
numsamall = length(unique(rawdata(2:end, col_samname))); %find numer of samples in run
[samlistall, indsamlistall] = unique(rawdata(2:end, col_samname)); %list of all samples in run and indecies of first row appearance
[comlistall, indcomlistall] = unique(rawdata(2:end, col_comname)); %creat list of all compounds and indecies of first row of compound in rawdata file
numcom = 0; numIS = 0; mcal = 0; maQC = 0; mmQC = 0; msam = 0; numNIS = 0; numallIS = 0; %reset counters

for i = 1:length(comlistall)

    if isequal(rawdata{indcomlistall(i) + 1, col_comtype}, 'Target Compound')
        numcom = numcom + 1;
        comlist(numcom, 1) = comlistall(i);
        commass(numcom, 1) = rawdata(indcomlistall(i) + 1, col_mzexp);
    else
        numallIS = numallIS + 1;
        allISlist(numallIS, 1) = comlistall(i);

        if find(strcmp(rawdata2(:, col_IS2), rawdata{indcomlistall(i) + 1, col_comname})) > 0
            numIS = numIS + 1;
            ISlist(numIS, 1) = comlistall(i);
            ISmass(numIS) = rawdata(indcomlistall(i) + 1, col_mzexp);
        else
            numNIS = numNIS + 1;
            NISlist(numNIS, 1) = comlistall(i);
            NISmass(numNIS) = rawdata(indcomlistall(i) + 1, col_mzexp);
        end

    end

end

for i = 1:length(samlistall)

    if isequal(rawdata{indsamlistall(i) + 1, col_samtype}, 'Cal Std') %if sample is a cal standard
        mcal = mcal + 1;
        callist{mcal, 1} = rawdata{indsamlistall(i) + 1, col_samname};
        % elseif contains(rawdata{indsamlistall(i)+1,col_samname},'CCV') || contains(rawdata{indsamlistall(i)+1,col_samname},'ISC') || contains(rawdata{indsamlistall(i)+1,col_samname},'LB') || contains(rawdata{indsamlistall(i)+1,col_samname},'ICV') || contains(rawdata{indsamlistall(i)+1,col_samname},'AFFF')|| contains(rawdata{indsamlistall(i)+1,col_samname},'ALE') %if an analytical QC
        %     maQC=maQC+1;
        %     aQClist{maQC,1} = rawdata{indsamlistall(i)+1,col_samname};
    elseif contains(rawdata{indsamlistall(i) + 1, col_samname}, 'LSC') || contains(rawdata{indsamlistall(i) + 1, col_samname}, 'X') || contains(rawdata{indsamlistall(i) + 1, col_samname}, 'MB') || contains(rawdata{indsamlistall(i) + 1, col_samname}, 'QC') %if a method QC
        mmQC = mmQC + 1;
        mQClist{mmQC, 1} = rawdata{indsamlistall(i) + 1, col_samname};
    elseif contains(rawdata{indsamlistall(i) + 1, col_samname}, 'DB') %if a double blank
    else
        msam = msam + 1;
        samlist{msam, 1} = rawdata{indsamlistall(i) + 1, col_samname};
    end

end

%nsam2=nsam;
%% 7) Extract Data
% Categories:
%   DB: DB
%   cal: calibration standards
%   aQC: CCV, ISC, LB (analytical QC samples)
%   mQC: LCS, LCSD, MB (method QC samples)
%   sam: samples
% Extract actual cal conc, actual IS conc, IS peak area, measured concentration

for i = 2:size(rawdata, 1) %for each row in rawdata file

    if isequal(rawdata{i, col_comtype}, 'Internal Standard') %if row is an internal standard
        comindex2 = find(strcmp(allISlist(:, 1), rawdata(i, col_comname))); %row index of compound

        if isequal(rawdata{i, col_samtype}, 'Cal Std')
            ncal2 = find(strcmp(callist(:, 1), rawdata(i, col_samname))); %row index of cal standard
            calactISconc2(comindex2, ncal2) = extractvalue(rawdata{i, col_conc}); %actual cal IS concentration
            calISPA2(comindex2, ncal2) = extractvalue(rawdata{i, col_comPA}); %cal IS Peak Area
            %elseif contains(rawdata{i,col_samname},'CCV') || contains(rawdata{i,col_samname},'ISC') || contains(rawdata{i,col_samname},'LB') || contains(rawdata{i,col_samname},'ICV') || contains(rawdata{i,col_samname},'AFFF') ||contains(rawdata{i,col_samname},'LSC') || contains(rawdata{i,col_samname},'X') || contains(rawdata{i,col_samname},'MB') || contains(rawdata{i,col_samname},'QC')||contains(rawdata{i,col_samname},'DB')
        else
            nsam2 = find(strcmp(samlist(:, 1), rawdata(i, col_samname)));
            ISPA(comindex2, nsam2) = extractvalue(rawdata{i, col_comPA});
            ISconc(comindex2, nsam2) = extractvalue(rawdata{i, col_conc});
        end

    else
        comindex = find(strcmp(comlist(:, 1), rawdata(i, col_comname))); %row index of compound

        if isequal(rawdata{i, col_samtype}, 'Cal Std') %if sample is a cal standard
            ncal = find(strcmp(callist(:, 1), rawdata(i, col_samname))); %row index of cal standard
            calname{comindex, ncal} = rawdata{i, col_samname}; %cal name
            calactconc(comindex, ncal) = extractvalue(rawdata{i, col_actconc}); %actual cal concentration
            calactISconc(comindex, ncal) = rawdata{i, col_ISact}; %actual cal IS concentration
            calISPA(comindex, ncal) = extractvalue(rawdata{i, col_ISPA}); %cal IS Peak Area
            calconc(comindex, ncal) = extractvalue(rawdata{i, col_conc}); %cal concentration
            calPA(comindex, ncal) = extractvalue(rawdata{i, col_comPA}); %peak area
            calISact(comindex, ncal) = extractvalue(rawdata{i, col_ISact});
            % elseif contains(rawdata{i,col_samname},'CCV') || contains(rawdata{i,col_samname},'ISC') || contains(rawdata{i,col_samname},'LB') || contains(rawdata{i,col_samname},'ICV') || contains(rawdata{i,col_samname},'AFFF') %|| contains(rawdata{i,col_samname},'QC') %if an analytical QC
            %     naQC = find(strcmp(aQClist(:,1),rawdata(i,col_samname))); %row index of aQC standard
            %     aQCname{comindex,naQC} = rawdata{i,col_samname}; %aQC name
            %     aQCactISconc(comindex,naQC) = rawdata{i,col_ISact}; %actual aQC IS concentration
            %     aQCISPA(comindex,naQC) = extractvalue(rawdata{i,col_ISPA}); %aQC IS Peak Area
            %     aQCconc(comindex,naQC) = extractvalue(rawdata{i,col_conc}); %aQC concentration
            %     aQCactconc(comindex,naQC) = extractvalue(rawdata{i,col_actconc});
            %     aQCPA(comindex,naQC) = extractvalue(rawdata{i,col_comPA}); %aQC peak area
        elseif contains(rawdata{i, col_samname}, 'LSC') || contains(rawdata{i, col_samname}, 'X') || contains(rawdata{i, col_samname}, 'MB') || contains(rawdata{i, col_samname}, 'QC') %if a method QC
            nmQC = find(strcmp(mQClist(:, 1), rawdata(i, col_samname))); %row index of mQC standard
            mQCname{comindex, nmQC} = rawdata{i, col_samname}; %mQC name
            mQCactISconc(comindex, nmQC) = rawdata{i, col_ISact}; %actual mQC IS concentration
            mQCISPA(comindex, nmQC) = extractvalue(rawdata{i, col_ISPA}); %mQC IS Peak Area
            mQCconc(comindex, nmQC) = extractvalue(rawdata{i, col_conc}); %cal concentration
            mQCPA(comindex, nmQC) = extractvalue(rawdata{i, col_comPA}); %mQC peak area
        elseif contains(rawdata{i, col_samname}, 'DB') %if a double blank
        else %if a sample
            nsam = find(strcmp(samlist(:, 1), rawdata(i, col_samname))); %row index of mQC standard
            samname{comindex, nsam} = rawdata{i, col_samname}; %sam name
            samactISconc(comindex, nsam) = rawdata{i, col_ISact}; %actual sample IS concentration
            samISPA(comindex, nsam) = extractvalue(rawdata{i, col_ISPA}); %aQC IS Peak Area
            samconc(comindex, nsam) = extractvalue(rawdata{i, col_conc}); %cal concentration
            samPA(comindex, nsam) = extractvalue(rawdata{i, col_comPA}); %sam peak area
        end

    end

end

%nsam=nsam2;
%% 8) Deturmine Calibration Range
[calsort, calorderindex] = sort(calactconc(1, :), 2); %find indicies of order
%reorder calconc matrix in order in increasing concentration
for i = 1:length(calorderindex)
    calorder(:, i) = calconc(:, calorderindex(i));
end

e = calorder ./ calsort;

for i = 1:size(e, 1)

    for j = 1:size(e, 2)

        if e(i, j) < 0.7 || e(i, j) > 1.3 || isnan(e(i, j))
            ee(i, j) = 0;
        else
            ee(i, j) = 1;
        end

    end

end

for i = 1:numcom %for each compound
    j = 0; f = 0;

    while f == 0
        j = j + 1;
        f = ee(i, j);

        if isequal(j, size(calsort, 2))
            f = 1;
        end

    end

    lowlim(i, 1) = calsort(j);

    if isequal(j, size(calsort, 2))
        lowlim(i, 1) = NaN;
    end

end

for i = 1:numcom %for each compound
    j = size(calsort, 2) + 1; f = 0;

    while f == 0
        j = j - 1;
        f = ee(i, j);

        if isequal(j, 1)
            f = 1;
        end

    end

    highlim(i, 1) = calsort(j);

    if isequal(j, 1)
        highlim(i, 1) = NaN;
    end

end

%% 9) Deturmine IS and NIS Peak Area Recovery
if isequal(matmethod, 'plasma') || isequal(matmethod, 'soil') || isequal(matmethod, 'dust')

    for i = 1:length(comlist)
        % Suche den Index des aktuellen internen Standards in der match list
        nis_ind = find(strcmp(rawdata2(:, col_com), comlist(i)), 1);

        %nis_ind = find(strcmp(rawdata2(:,col_IS2), ISlist(i)),1); % Index des zugehörigen NIS
        nis_name = rawdata2(nis_ind, col_NIS);
        nisfin_ind = find(strcmp(allISlist(:, 1), nis_name));

        if ~isempty(nis_ind) % Wenn ein passender NIS gefunden wurde
            % Berechne die IS-Wiederfindung für den aktuellen internen Standard
            samISrec(i, :) = samISPA(i, :) ./ ISPA(nisfin_ind, :) ./ nanmean(calISPA(i, :) ./ calISPA2(nisfin_ind, :), 2) .* (mean(calactISconc(i, :), 2) ./ samactISconc(i, :));

        else
            warning(['Kein NIS für den internen Standard ' char(ISlist(i)) ' gefunden.']);
        end

    end

    for i = 1:length(NISlist)
        nis_name = rawdata2(i + 1, col_NIS);
        nisfin_ind = find(strcmp(allISlist(:, 1), nis_name));
        NISrec(i, :) = ISPA(nisfin_ind, :) ./ nanmean(calISPA2(nisfin_ind, :), 2);
    end

    for i = 1:length(ISlist)
        is_ind = find(strcmp(rawdata2(:, col_IS2), ISlist(i)), 1); % Index des zugehörigen NIS
        is2_ind = find(strcmp(allISlist(:, 1), ISlist(i)), 1);
        nis_name = rawdata2(is_ind, col_NIS);
        nisfin_ind = find(strcmp(allISlist(:, 1), nis_name));
        ISrec(i, :) = ISPA(is2_ind, :) ./ ISPA(nisfin_ind, :) ./ nanmean(calISPA2(is2_ind, :) ./ calISPA2(nisfin_ind, :), 2) .* (mean(calactISconc2(is2_ind, :), 2) ./ ISconc(is2_ind, :));
    end

    % m2pfoa_ind = find(ISlist=="M2PFOA"); % Index of M2PFOA Peak Areas
    % samISrec = samISPA./samPA(m2pfoa_ind,:)./nanmean(calISPA./calPA(m2pfoa_ind,:),2).*(mean(calactISconc,2)./samactISconc); %Calculate sample IS recovery
    % M2PFOArec = samPA(m2pfoa_ind,:)./nanmean(calPA(m2pfoa_ind,:),2);
else
    samISrec = samISPA ./ mean(calISPA, 2, "omitnan") .* (mean(calactISconc, 2) ./ samactISconc); %calculate sample IS recovery taking into account any difference among IS actual conc values
    ISrec = ISPA ./ mean(calISPA2, 2, "omitnan") .* (mean(calactISconc2, 2) ./ ISconc);
end

%% 10) Check LCMS Analytical Accuracy
% Calculate CCV recoveries
% CCVindex = find(contains(aQCname(1,:),'CCV')); %find the indicies of the CCVs
% CCVrec = aQCconc(:,CCVindex)./aQCactconc(:,CCVindex);%CCVactconc;
% CCVname = aQCname(:,CCVindex);

% % Calculate ISC recoveries
% ISCindex = find(contains(aQCname(1,:),'ISC')); %find the indicies of the ISCs
% ISCrec = aQCconc(:,ISCindex)./aQCactconc(:,ISCindex);%ISCactconc;
% ISCname = aQCname(:,ISCindex);

% % Calculate LB concentrations
% LBindex = find(contains(aQCname(1,:),'LB')); %find the indicies of the LBs
% LBconc = aQCconc(:,LBindex);
% LBname = aQCname(:,LBindex);

% % Calculate EPA Bullseye concentrations
% ICVindex = find(contains(aQCname(1,:),'ICV')); %find the indicies of the ICVs
% ICVrec = aQCconc(:,ICVindex)./ICVactconc;
% ICVname = aQCname(:,ICVindex);
%
% aQCall = [LBconc];
% aQCallname = {LBname{1,:}};

%% 11) Check Method Accuracy and Deturmine Reporting Limit
if mmQC > 0 %if there are method QC samples

    % Calculate LCS and LLLCS concentrations
    LCSindex = find(contains(mQCname(1, :), 'X')); %find the indicies of the LCS and LCSDs
    LCSrec = mQCconc(:, LCSindex); %/aQCactconc(:,LCSindex); %LCSactconc;
    LCSname = mQCname(:, LCSindex);

    % Calculate MB concentrations
    MBindex = find(contains(mQCname(1, :), 'MB')); %find the indicies of the LBs
    MBconc = mQCconc(:, MBindex);
    MBname = mQCname(:, MBindex);

    %Identify Reporting Limit
    if max(MBindex) > 0 %if there is a MB sample present

        for i = 1:numcom %for each compound

            if max(MBconc(i, :)) > lowlim(i) %if the maximum method blank concentration is higher than the lower calibration limit
                RL(i, 1) = max(MBconc(i, :)); %the reporting limit is the maximum method blank concentration
            else
                RL(i, 1) = lowlim(i); %otherwise the reporting limit is the lower calibration limit
            end

        end

        if max(LCSindex) > 0 %if there is a LCS/LCSD sample present
            mQCall(:, :) = [LCSrec(:, :) MBconc(:, :)]; %combined matrix of LCS/LCSD and MB
        else
            mQCall(:, :) = MBconc(:, :); % matrix of only MB
        end

    else
        RL(i, 1) = lowlim(i); %otherwise the reporting limit is the lower calibration limit

        if max(LCSindex) > 0 %if there is a LCS/LCSD sample present
            mQCall(:, :) = LCSrec(:, :); %Matrix of only LCS/LCSD
        end

    end

else

    for i = 1:numcom %for each compound
        RL(i, 1) = lowlim(i); %otherwise the reporting limit is the lower calibration limit
    end

end

%% 12) Calculate Matrix Conc, Matrix LoQ Range, and Replace Values Outside of Matrix LoQ
%if isequal(matmethod,'plasma')
%    factor = 100/10^6*2/100*10^3; %In-vial conc  * 2 * 100uL *1L/10^6uL / 100uL * 10^3uL/mL  = plasma concentration [ng/mL]
%    units='[ng/mL]';
%elseif isequal(matmethod,'soil')
%    factor = 100/10^6*2/100*10^3; %In-vial conc  * 2 * 100uL *1L/10^6uL / 100uL * 10^3uL/mL  = plasma concentration [ng/mL]
%    units='[ng/g]';
%elseif isequal(matmethod,'dust')
%    factor = 100/10^6*2/100*10^3; %In-vial conc  * 2 * 100uL *1L/10^6uL / 100uL * 10^3uL/mL  = plasma concentration [ng/mL]
%    units='[ng/g]';
%elseif isequal(matmethod,'aqueous')
%   factor = 100/10^6*2/100*10^3; %In-vial conc  * 2 * 100uL *1L/10^6uL / 100uL * 10^3uL/mL  = plasma concentration [ng/mL]
%   units='[ng/L]';
%end
%matconc = samconc * factor;

%% 12) Calculate Matrix Conc, Matrix LoQ Range, and Replace Values Outside of Matrix LoQ
if isequal(matmethod, 'soil') || isequal(matmethod, 'dust') %if soil or dust method
    %[soildata,soildataind] = xlsread(strcat(path_name,'samplemass.xlsx'),file_name); %read in matrix mass data
    matmass = 1; %(soildata(:,1).*(1-soildata(:,2)))'; %extract matrix mass data
end

for i = 1:msam %for each sample
    % Deturmine matrix extraction factor and typical mass extracted
    if isequal(matmethod, 'soil') %if matrix is soil
        matexfactor = 1; %0.4/0.3 * 1.5/1000; %total vial vol (0.4mL) / extract vial vol (0.1mL) * total extract volume (1.5mL) / 1000 [mL/L] = units [L]
    elseif isequal(matmethod, 'dust') %otherwise if matrix is dust
        matexfactor = 0.4/0.12 * 1.5/1000 * 1.5/0.5; %total vial vol (0.4mL) / extract vial vol (0.12mL) * SPE extract volume (1.5mL) / 1000 [mL/L] * total extract (1.5mL) / SPE extract (0.5mL) = units [L]
    elseif isequal(matmethod, 'plasma')
        matexfactor = 100/10 ^ 6 * 2/100 * 10 ^ 3; %In-vial conc  * 2 * 100uL *1L/10^6uL / 100uL * 10^3uL/mL = plasma concentration [ng/mL]
    elseif isequal(matmethod, 'aqueous')
        matexfactor = 1; %In-vial conc * 1.5mL volume / 0.9 mL water = concentration in water sample [ng/L]
    end

    % Calculate matrix concentration [ng/g] and matrix upper & lower LoQ range
    sammass(:, i) = samconc(:, i) * matexfactor; %total sample mass [ng] = vial conc [ng/L] * matrix extraction factor [L]

    if isequal(matmethod, 'soil') || isequal(matmethod, 'dust')
        matsamconc(:, i) = sammass(:, i) ./ matmass; %(:,i); %matrix concentration [ng/g] = sample mass [ng] / matrix mass [g]
        mhighlim(:, i) = highlim * matexfactor / matmass; %(i); %upper quant limit [ng/g] = upper caliration limit [ng/L] * mat extraction factor [L] / mass extracted [g]
        mlowlim(:, i) = RL * matexfactor / matmass; %(i); %upper quant limit [ng/g] = reporting limit [ng/L] * mat extraction factor [L] / mass extracted [g]
    else
        matsamconc(:, i) = sammass(:, i); %matrix concentration = in-vial concentration * matrix extraction factor [mass/volume]
        mhighlim(:, i) = highlim * matexfactor; %upper quant limit [ng/g] = upper caliration limit [ng/L] * mat extraction factor [L] / mass extracted [g]
        mlowlim(:, i) = RL * matexfactor; %upper quant limit [ng/g] = reporting limit [ng/L] * mat extraction factor [L] / mass extracted [g]
    end

    % Replace Values outside of the matrix LoQ
    for j = 1:numcom %for each compound

        if matsamconc(j, i) > mhighlim(j) %if matrix conc is higher than matrix LoQ
            matconcLoQ{j, i} = matsamconc(j, i); % matconcLoQ{j,i} = sprintf('>%.2f', mhighlim(j, i)); %replace concentration with '> upper LoQ #'
        elseif matsamconc(j, i) < mlowlim(j) || isnan(matsamconc(j, i)) %if matrix conc is lower than matrix LoQ
            matconcLoQ{j, i} = sprintf('<%.2f', mlowlim(j, i)); %replace concentration with '> lower LoQ #'
        else
            matconcLoQ{j, i} = matsamconc(j, i); %matrix concentraiton is within LoQ, no change necessary
        end

    end

end

% % 13) Correct Diluted Values
% for i=1:nsam %for each sample
%     if contains(samname(i),'_d') && contains(samname(i),'x') %if sample is diluted (i.e. has '_d##x' in name)
%         dil=extractBetween(samname(i),'_d','x'); ndil=str2num(dil{:}); %dilution factor is the number between the '_d' and 'x'
%         for j=1:numcom %for each compound
%             if contains(num2str(matconcLoQ{j,i}),'>') %if the compound is above LoQ
%                 dmatconcLoQ{j,i}=sprintf('>%.2f',mhighlim(j,i)*ndil); %recalculate the LoQ based on the dilution factor
%             elseif contains(num2str(matconcLoQ{j,i}),'<') %if compound is below the LoQ
%                 dmatconcLoQ{j,i}=sprintf('<%.2f',mlowlim(j,i)*ndil); %recalculate the LoQ based on dilution factor
%             else %otherwise
%                 dmatconcLoQ{j,i}=cell2mat(matconcLoQ(j,i))*ndil; %multiple by dilution factor
%             end
%         end
%     else %if not diluted
%         dmatconcLoQ(:,i)=matconcLoQ(:,i);
%     end
% end

%% 14) Write .xlsx Data File for Full Results
% All Sample Matrix Concentrations
tab1 = {}; %define tab as cell array
tab1(1, 1) = {'Compound'}; %define cell A1
tab1(2:numcom + 1, 1) = comlist(:, 1); %define compounds beginning in cell A2 extending vertically
tab1(1, 2) = {'Precursor Mass'}; %define cell B1
tab1(2:end, 2) = commass(:, 1); %fill compound masses in B2 extending vertically
tab1(1, 3:msam + 2) = samname(1, :); %define samples names in cells D1 extending horizontantlly
tab1(2:end, 3:end) = matconcLoQ(:, :); %fill sample concentrations extending vertically (per sample) and horizontally for each sample

writecell(tab1, fullfile(path_name, strcat(file_name, '_resultsnis.xlsx')), 'Sheet', 'Sample Conc'); %write tab1 into results xlsx file

% Sample IS Recovery
tab2 = {}; %define tab as cell array
tab2(1, 1) = {'Surrogate Recovery [%]'}; %define cell A1
tab2(2:numcom + 1, 1) = comlist(:, 1);
tab2(1, 2:msam + 1) = samname(1, :);
tab2(2:end, 2:end) = num2cell(samISrec);

if isequal(matmethod, 'soil') || isequal(matmethod, 'dust') || isequal(matmethod, 'plasma')
    tab2(numcom + 3, 1) = {'NIS recovery'};
    tab2(numcom + 4:numcom + 10, 1) = NISlist(:, 1);
    tab2(numcom + 4:numcom + 10, 2:end) = num2cell(NISrec(:, :));
end

writecell(tab2, fullfile(path_name, strcat(file_name, '_resultsnis.xlsx')), 'Sheet', 'Surrogate Recovery');

% Raw Matrix Concentrations [ng/L] of [ng/g]
tab3 = {}; %define tab as cell array
tab3(1, 1) = {strcat('Measured Matrix Concentration [', units, '] (NO DILUTION CORRECTIONS)')}; %define cell A1
tab3(2:numcom + 1, 1) = comlist(:, 1);
tab3(1, 2:size(samname, 2) + 1) = samname(1, :);
tab3(2:end, 2:end) = num2cell(matsamconc);

writecell(tab3, fullfile(path_name, strcat(file_name, '_resultsnis.xlsx')), 'Sheet', 'Raw Corrected Conc');

toc
%% color data in surrogate recovery
%
% % Specify the file name
% excelFileName = fullfile(path_name,strcat(file_name, '_results.xlsx'));
% sheetName = 'Surrogate Recovery';
%
% % Read the data from the Excel sheet
% data = readtable(excelFileName, 'Sheet', sheetName);
% outOfRange= false(size(data));
% outOfRange2= false(size(data));
% % Identify which values are outside the desired range (0.5 to 1.5)
% outOfRange(:,2:end) = (data{:,2:end} < 0.5) | (data{:,2:end} > 1.5); % Adjust this according to the structure of your table
% outOfRange2(:,2:end) = (data{:,2:end} < 0.3) | (data{:,2:end} > 2); % Adjust this according to the structure of your table
%
% % Only proceed if there are out-of-range values
% if any(outOfRange, 'all')
%     % Start an instance of Excel
%     Excel = actxserver('Excel.Application');
%     Workbook = Excel.Workbooks.Open(fullfile(excelFileName));
%     Sheet = Workbook.Sheets.Item(sheetName);
%     Excel.Visible = false; % Set this to false if you don't want Excel to open visibly
%
%     % Loop over the cells to change colors
%     [rows, cols] = find(outOfRange);
%     % Loop over the cells to change colors
%     for i = 1:length(rows)
%         rowIndex = rows(i) + 1; % Adjust for Excel's 1-based indexing
%         colIndex = cols(i); % Assuming first column is headers and data starts from the second column
%
%         % Construct the Excel cell reference (e.g., "B2")
%         cellRef = [char(64 + colIndex) num2str(rowIndex)]; % Converts to A1, B2 notation
%
%         % Now get the range using the cell reference
%         Cell = Sheet.Range(cellRef);
%
%         % Change the cell color to yellow for out-of-range values
%         Cell.Interior.Color = 65535; % Set the cell color to yellow
%     end
%        % Loop over the cells to change colors
%     [rows, cols] = find(outOfRange2);
%     % Loop over the cells to change colors
%     for i = 1:length(rows)
%         rowIndex = rows(i) + 1; % Adjust for Excel's 1-based indexing
%         colIndex = cols(i); % Assuming first column is headers and data starts from the second column
%
%         % Construct the Excel cell reference (e.g., "B2")
%         cellRef = [char(64 + colIndex) num2str(rowIndex)]; % Converts to A1, B2 notation
%
%         % Now get the range using the cell reference
%         Cell = Sheet.Range(cellRef);
%
%         % Change the cell color to yellow for out-of-range values
%         Cell.Interior.Color = 255; % Set the cell color to red
%     end
%     % Save and close the workbook
%     Workbook.Save;
%     Workbook.Close;
%     Excel.Quit;
%     % Clean up the COM server
%     delete(Excel);
% end
%
%

%% Functions
function value = extractvalue(input)

    if ischar(input)
        value = NaN;
    elseif isequal(input, [])
        value = NaN;
    else
        value = input;
    end

end
