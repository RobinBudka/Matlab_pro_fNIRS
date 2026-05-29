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
plot(timestamps, Raw_Signals(9,:) - Raw_Signals(9,Mrk_Locs(2)),'linewidth',1, 'Color', 'B');
hold on
plot(timestamps, Raw_Signals(10,:) - Raw_Signals(10,Mrk_Locs(2)),'linewidth',1, 'Color', 'R');
title('Ch1')
ylim([min(Raw_Signals(9,:) - Raw_Signals(9,Mrk_Locs(1))) max(Raw_Signals(9,:) - Raw_Signals(9,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,end))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')
title('Raw fNIRS signal')

subplot(1,3,2)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Wavelet_Signals(9,:) - Wavelet_Signals(9,Mrk_Locs(2)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Wavelet_Signals(10,:) - Wavelet_Signals(10,Mrk_Locs(2)),'linewidth',1, 'Color', 'B');
title('Ch1')
ylim([min(Wavelet_Signals(9,:) - Wavelet_Signals(9,Mrk_Locs(1))) max(Wavelet_Signals(9,:) - Wavelet_Signals(9,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,end))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')
title('Filtered with wavelet')

subplot(1,3,3)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Poly_Signals(9,:) - Poly_Signals(9,Mrk_Locs(2)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Poly_Signals(10,:) - Poly_Signals(10,Mrk_Locs(2)),'linewidth',1, 'Color', 'B');
title('Ch1')
ylim([min(Poly_Signals(9,:) - Poly_Signals(9,Mrk_Locs(1))) max(Poly_Signals(9,:) - Poly_Signals(9,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,end))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')
title('Filtered with wavelet and polynomial')


%% Filtered signals
[c, r] = size(Poly_Signals);
pocet_grafu = c / 2;
ymax = zeros(1, pocet_grafu);
ymin = zeros(1, pocet_grafu);

for n = 1:pocet_grafu
    sig1 = 2*n - 1;
    sig2 = 2*n;

    shifted_sig1 = Poly_Signals(sig1, :) - Poly_Signals(sig1, Mrk_Locs(1));
    shifted_sig2 = Poly_Signals(sig2, :) - Poly_Signals(sig2, Mrk_Locs(1));
    
    max_sig1 = max(shifted_sig1);
    max_sig2 = max(shifted_sig2);
    ymax(n) = max([max_sig1, max_sig2]);
    
    min_sig1 = min(shifted_sig1);
    min_sig2 = min(shifted_sig2);
    ymin(n) = min([min_sig1, min_sig2]);
end

figure(Name = 'Filtered trend pairs, Red = HbO, Blue = HbR');
% --Right side--
subplot(2,4,3)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Poly_Signals(1,:) - Poly_Signals(1,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Poly_Signals(2,:) - Poly_Signals(2,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch1')
ylim([ymin(1) ymax(1)]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,end))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

subplot(2,4,4)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Poly_Signals(3,:) - Poly_Signals(3,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Poly_Signals(4,:) - Poly_Signals(4,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch2')
ylim([ymin(2) ymax(2)]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,end))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

subplot(2,4,8)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Poly_Signals(5,:) - Poly_Signals(5,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Poly_Signals(6,:) - Poly_Signals(6,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch3')
ylim([ymin(3) ymax(3)]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,end))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

subplot(2,4,7)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Poly_Signals(7,:) - Poly_Signals(7,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Poly_Signals(8,:) - Poly_Signals(8,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch4')
ylim([ymin(4) ymax(4)]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,end))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

% --Left side--
subplot(2,4,2)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Poly_Signals(9,:) - Poly_Signals(9,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Poly_Signals(10,:) - Poly_Signals(10,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch5')
ylim([ymin(5) ymax(5)]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,end))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

subplot(2,4,1)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Poly_Signals(11,:) - Poly_Signals(11,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Poly_Signals(12,:) - Poly_Signals(12,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch6')
ylim([ymin(6) ymax(6)]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,end))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

subplot(2,4,5)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Poly_Signals(13,:) - Poly_Signals(13,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Poly_Signals(14,:) - Poly_Signals(14,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch7')
ylim([ymin(7) ymax(7)]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,end))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')

subplot(2,4,6)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Poly_Signals(15,:) - Poly_Signals(15,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Poly_Signals(16,:) - Poly_Signals(16,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch8')
ylim([min(Poly_Signals(15,:) - Poly_Signals(15,Mrk_Locs(1))) max(Poly_Signals(15,:) - Poly_Signals(15,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,end))])
hold off
ylabel('Amplitude [V]')
xlabel('Time [s]')
end