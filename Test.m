clc; clear all; close all;
%  This function serves for diagnosis and testing of filters and other
%  methods. Should be clean and faster to use and adapt then full filtering
%  and showing functions.

%% File selection (FILE EXPLORER)

[file, path] = uigetfile('*.mat', 'Choose data file (.mat)');

if isequal(file, 0)
    disp('File selection stopped');
    return;
end

% Loading data from selected file
full_path = fullfile(path, file);
disp(['Loading data from file: ', file]);
load(full_path);


%% Marker Editor

Marker = Out(35,1000:end);

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

%% Signal Extraction

timestamps = Out(1,1000:end);
SampleFreq = 1/(timestamps (1,2) - timestamps(1,1))
HR = Out(34,1000:end);
n=10; %wavelet order
w='db15'; % signal wavelet

%% Filters
Trends = zeros(16, length(timestamps));

for i = 1:16
    raw_signal = Out(i+1, 1000:end);
    
    wavelet_signal = wden(raw_signal, 'rigrsure', 's', 'one', n, w);
    
    % clean_signal = filloutliers(wavelet_signal, 'spline', 'movmedian', 250);

    Trends(i, :) = wavelet_signal;
end

% Plots
figure();
plot(timestamps, wden(Out(2,1000:end), 'rigrsure', 's', 'one', n, w)-Out(2,Mrk_Locs(2)))
hold on
plot(timestamps, wden(Out(3,1000:end), 'rigrsure', 's', 'one', n, w)-Out(3,Mrk_Locs(2)))
hold on
plot(timestamps, Marker*0.001)
hold off

% figure(Name = 'Signal processing stages');
% 
% subplot(3,1,1)
% plot(timestamps, raw_signal, 'linewidth',1, 'Color', 'R');
% hold on
% plot(timestamps, Marker, 'linewidth',1, 'Color', 'M')
% hold off
% title('Raw signal')
% 
% subplot(3,1,2)
% plot(timestamps, wavelet_signal, 'linewidth',1, 'Color', 'R');
% hold on
% plot(timestamps, Marker, 'linewidth',1, 'Color', 'M')
% hold off
% title('Signal after wavelet')
% 
% subplot(3,1,3)
% plot(timestamps, clean_signal, 'linewidth',1, 'Color', 'R');
% hold on
% plot(timestamps, Marker, 'linewidth',1, 'Color', 'M')
% hold off
% title('Signal after wavelet and fill outliers')

%% Signal quality
SQLvs = zeros(16, length(timestamps));

for i = 35: 50
    SQ_Lev = Out(i, 1000:end);
    SQLvs(i-34,:) = SQ_Lev;  
end

figure(Name = 'SQ level');
plot(timestamps, SQLvs(1,:)*0.001)
hold on
plot(timestamps, wden(Out(3,1000:end), 'rigrsure', 's', 'one', n, w)-Out(3,Mrk_Locs(2)))

SQ = zeros(8, length(timestamps));

for i = 51: 58
    Sig_Qual = Out(i, 1000:end);
    SQ(i-50,:) = Sig_Qual;  
end

figure(Name = 'SQ');
plot(timestamps, SQ(2,:))
