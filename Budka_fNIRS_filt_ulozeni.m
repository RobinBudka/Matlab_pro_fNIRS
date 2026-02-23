clc; clear all; close all;

% --- 1. File selection (FILE EXPLORER) ---
[file, path] = uigetfile('*.mat', 'Choose data file (.mat)');

if isequal(file, 0)
    disp('File selection stopped');
    return;
end

% Loading data from selected file
full_path = fullfile(path, file);
disp(['Loading data from file: ', file]);
load(full_path);

% Excel file name
[~, name_only, ~] = fileparts(file);
excel_name_short = [name_only, '.xlsx']; 
full_path_to_excel = fullfile(path, excel_name_short);

timestamps = Out(1,1000:end);
SampleFreq = 1/(timestamps (1,2) - timestamps(1,1))

Marker = Out(59,1000:end);

if length(Marker) > 300
    Marker(end-300) = 1;
end

Mrk_Locs = find(Marker == 1);



if length (Mrk_Locs) > 4

    i = 1;
    while i < length(Mrk_Locs)
        Mrk_Diff = Mrk_Locs(i+1) - Mrk_Locs(i);

        if Mrk_Diff < 2000
            Marker(Mrk_Locs(i)) = 0;
            Mrk_Locs(i) = [];
        else
            i = i + 1;
        end
    end

else
    disp('Please check Markers and add manually.')
end

HR = Out(34,1000:end);

n=10; %wavelet order
w='db15'; % signal wavelet

HbO1=Out(2,1000:end);
trend_HbO1=wden(HbO1,'rigrsure','s','one',n,w);


HbR1=Out(3,1000:end);
trend_HbR1=wden(HbR1,'rigrsure','s','one',n,w);


HbO2=Out(4,1000:end)-0.047;
trend_HbO2=wden(HbO2,'rigrsure','s','one',n,w);

HbR2=Out(5,1000:end);
trend_HbR2=wden(HbR2,'rigrsure','s','one',n,w);

HbO3=Out(6,1000:end);
trend_HbO3=wden(HbO3,'rigrsure','s','one',n,w);

HbR3=Out(7,1000:end);
trend_HbR3=wden(HbR3,'rigrsure','s','one',n,w);

HbO4=Out(8,1000:end);
trend_HbO4=wden(HbO4,'rigrsure','s','one',n,w);

HbR4=Out(9,1000:end);
trend_HbR4=wden(HbR4,'rigrsure','s','one',n,w);

HbO5=Out(10,1000:end);
trend_HbO5=wden(HbO5,'rigrsure','s','one',n,w);

HbR5=Out(11,1000:end);
trend_HbR5=wden(HbR5,'rigrsure','s','one',n,w);

HbO6=Out(12,1000:end);
trend_HbO6=wden(HbO6,'rigrsure','s','one',n,w);

HbR6=Out(13,1000:end);
trend_HbR6=wden(HbR6,'rigrsure','s','one',n,w);

HbO7=Out(14,1000:end);
trend_HbO7=wden(HbO7,'rigrsure','s','one',n,w);

HbR7=Out(15,1000:end);
trend_HbR7=wden(HbR7,'rigrsure','s','one',n,w);

HbO8=Out(16,1000:end);
trend_HbO8=wden(HbO8,'rigrsure','s','one',n,w);

HbR8=Out(17,1000:end);
trend_HbR8=wden(HbR8,'rigrsure','s','one',n,w);

Trends = [trend_HbO1; trend_HbR1; trend_HbO2; trend_HbR2; trend_HbO3; trend_HbR3; trend_HbO4; trend_HbR4; ...
    trend_HbO5; trend_HbR5; trend_HbO6; trend_HbR6; trend_HbO7; trend_HbR7; trend_HbO8; trend_HbR8];

figureShower (Trends,timestamps, Marker, Mrk_Locs, HR)

