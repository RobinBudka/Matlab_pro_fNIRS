clc; clear all; close all;

% This is the final program main backbone, only adapt when new parts work
% properly. Keep optimized without unnecessary repeating.

%% File selection (FILE EXPLORER)

[file, path] = uigetfile('*.mat', 'Choose data file (.mat)');

if isequal(file, 0)
    disp('File selection stopped');
    return;
end

%% Loading data from selected file
full_path = fullfile(path, file);
disp(['Loading data from file: ', file]);
load(full_path);

%% Excel file name

[~, name_only, ~] = fileparts(file);
excel_name_short = [name_only, '.xlsx']; 
full_path_to_excel = fullfile(path, excel_name_short);

%% Marker Editor

Marker = Out(59,1000:end);

if length(Marker) > 300
    Marker(end-300) = 1;
    disp('Marker added 300 samples from end')
end

Mrk_Locs = find(Marker == 1);

if length (Mrk_Locs) > 4

    i = 1;
    while i < length(Mrk_Locs)
        Mrk_Diff = Mrk_Locs(i+1) - Mrk_Locs(i);

        if Mrk_Diff < 2000
            Marker(Mrk_Locs(i)) = 0;
            Mrk_Locs(i) = [];
            disp('Marker removed')
        else
            i = i + 1;
        end
    end

else
    disp('Please check Markers and add manually.')
end

%% Signal extraction

timestamps = Out(1,1000:end);
SampleFreq = 1/(timestamps (1,2) - timestamps(1,1))

HR = Out(34,1000:end);
Raw_Signals(1:16, :) = Out(2:17, 1000:end);

[Wavelet_Signals, Poly_Signals] = Fn_Filter(Raw_Signals, timestamps);
Fn_FigShower (Raw_Signals, Wavelet_Signals, Poly_Signals, timestamps, Marker, Mrk_Locs, HR);
% Fn_save2excel (Trends, Marker, timestamps, full_path_to_excel);
