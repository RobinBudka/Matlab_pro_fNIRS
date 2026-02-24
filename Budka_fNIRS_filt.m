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

Trends = zeros(16, length(timestamps));

for i = 1:16
    raw_signal = Out(i+1, 1000:end);
    
    wavelet_signal = wden(raw_signal, 'rigrsure', 's', 'one', n, w);
    
    clean_signal = filloutliers(wavelet_signal, 'spline', 'movmedian', 250);

    Trends(i, :) = clean_signal;
end

figure(Name = 'Signal processing stages');

subplot(3,1,1)
plot(timestamps, raw_signal, 'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Marker, 'linewidth',1, 'Color', 'M')
hold off
title('Raw signal')

subplot(3,1,2)
plot(timestamps, wavelet_signal, 'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Marker, 'linewidth',1, 'Color', 'M')
hold off
title('Signal after wavelet')

subplot(3,1,3)
plot(timestamps, clean_signal, 'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Marker, 'linewidth',1, 'Color', 'M')
hold off
title('Signal after wavelet and fill outliers')


figureShower (Trends,timestamps, Marker, Mrk_Locs, HR)
