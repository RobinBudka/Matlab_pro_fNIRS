function [] = Fn_FigShower(Raw_Signals, Wavelet_Signals, Poly_Signals, timestamps, Marker, Mrk_Locs, HR)
%This function is responsible for receiving data and drawing that data to
%figures

Marker = Marker - 0.5;
%% Heart rate
figure(Name = 'Heart Rate');
plot(timestamps, HR, 'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Marker*90, 'linewidth',1, 'Color', 'M')
hold off
ylabel('Heart rate [BPM]')
xlabel('Time [s]')

%% Filter comparison
figure();
subplot(1,3,1)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Raw_Signals(1,:) - Raw_Signals(1,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Raw_Signals(2,:) - Raw_Signals(2,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch1')
ylim([min(Raw_Signals(1,:) - Raw_Signals(1,Mrk_Locs(1))) max(Raw_Signals(1,:) - Raw_Signals(1,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')
title('Raw fNIRS signal')

subplot(1,3,2)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Wavelet_Signals(1,:) - Wavelet_Signals(1,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Wavelet_Signals(2,:) - Wavelet_Signals(2,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch1')
ylim([min(Wavelet_Signals(1,:) - Wavelet_Signals(1,Mrk_Locs(1))) max(Wavelet_Signals(1,:) - Wavelet_Signals(1,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')
title('Filtered with wavelet')

subplot(1,3,3)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Poly_Signals(1,:) - Poly_Signals(1,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Poly_Signals(2,:) - Poly_Signals(2,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch1')
ylim([min(Poly_Signals(1,:) - Poly_Signals(1,Mrk_Locs(1))) max(Poly_Signals(1,:) - Poly_Signals(1,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')
title('Filtered with wavelet and polynomial')


%% Filtered signals
figure(Name = 'Filtered trend pairs, Red = HbO, Blue = HbR');
% --Right side--
subplot(2,4,3)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Raw_Signals(1,:) - Raw_Signals(1,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Raw_Signals(2,:) - Raw_Signals(2,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch1')
ylim([min(Raw_Signals(1,:) - Raw_Signals(1,Mrk_Locs(1))) max(Raw_Signals(1,:) - Raw_Signals(1,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

subplot(2,4,4)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Raw_Signals(3,:) - Raw_Signals(3,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Raw_Signals(4,:) - Raw_Signals(4,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch2')
ylim([min(Raw_Signals(3,:) - Raw_Signals(3,Mrk_Locs(1))) max(Raw_Signals(3,:) - Raw_Signals(3,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

subplot(2,4,8)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Raw_Signals(5,:) - Raw_Signals(5,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Raw_Signals(6,:) - Raw_Signals(6,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch3')
ylim([min(Raw_Signals(5,:) - Raw_Signals(5,Mrk_Locs(1))) max(Raw_Signals(5,:) - Raw_Signals(5,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

subplot(2,4,7)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Raw_Signals(7,:) - Raw_Signals(7,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Raw_Signals(8,:) - Raw_Signals(8,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch4')
ylim([min(Raw_Signals(7,:) - Raw_Signals(7,Mrk_Locs(1))) max(Raw_Signals(7,:) - Raw_Signals(7,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

% --Left side--
subplot(2,4,2)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Raw_Signals(9,:) - Raw_Signals(9,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Raw_Signals(10,:) - Raw_Signals(10,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch5')
ylim([min(Raw_Signals(9,:) - Raw_Signals(9,Mrk_Locs(1))) max(Raw_Signals(9,:) - Raw_Signals(9,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

subplot(2,4,1)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Raw_Signals(11,:) - Raw_Signals(11,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Raw_Signals(12,:) - Raw_Signals(12,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch6')
ylim([min(Raw_Signals(11,:) - Raw_Signals(11,Mrk_Locs(1))) max(Raw_Signals(11,:) - Raw_Signals(11,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

subplot(2,4,5)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Raw_Signals(13,:) - Raw_Signals(13,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Raw_Signals(14,:) - Raw_Signals(14,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch7')
ylim([min(Raw_Signals(13,:) - Raw_Signals(13,Mrk_Locs(1))) max(Raw_Signals(13,:) - Raw_Signals(13,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

subplot(2,4,6)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Raw_Signals(15,:) - Raw_Signals(15,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Raw_Signals(16,:) - Raw_Signals(16,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch8')
ylim([min(Raw_Signals(15,:) - Raw_Signals(15,Mrk_Locs(1))) max(Raw_Signals(15,:) - Raw_Signals(15,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')
end