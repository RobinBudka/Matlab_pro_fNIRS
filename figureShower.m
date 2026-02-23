function [] = figureShower(Trends,timestamps, Marker, Mrk_Locs, HR)
%This function is responsible for receiving data and drawing that data to
%figures

Marker = Marker - 0.5;

figure(Name = 'Heart Rate');
plot(timestamps, HR, 'linewidth',1, 'Color', 'R');

figure(Name = 'All trends');
plot(timestamps, Trends,'linewidth',1);
hold on
plot(timestamps, Marker,'linewidth',1);

figure(Name = 'Trend pairs, Red = HbO, Blue = HbR');
% --Right side--
subplot(2,4,3)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Trends(1,:) - Trends(1,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Trends(2,:) - Trends(2,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch1')
ylim([min(Trends(1,:) - Trends(1,Mrk_Locs(1))) max(Trends(1,:) - Trends(1,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off

subplot(2,4,4)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Trends(3,:) - Trends(3,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Trends(4,:) - Trends(4,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch2')
ylim([min(Trends(3,:) - Trends(3,Mrk_Locs(1))) max(Trends(3,:) - Trends(3,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off

subplot(2,4,8)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Trends(5,:) - Trends(5,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Trends(6,:) - Trends(6,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch3')
ylim([min(Trends(5,:) - Trends(5,Mrk_Locs(1))) max(Trends(5,:) - Trends(5,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off

subplot(2,4,7)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Trends(7,:) - Trends(7,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Trends(8,:) - Trends(8,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch4')
ylim([min(Trends(7,:) - Trends(7,Mrk_Locs(1))) max(Trends(7,:) - Trends(7,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off

% --Left side--
subplot(2,4,2)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Trends(9,:) - Trends(9,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Trends(10,:) - Trends(10,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch5')
ylim([min(Trends(9,:) - Trends(9,Mrk_Locs(1))) max(Trends(9,:) - Trends(9,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off

subplot(2,4,1)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Trends(11,:) - Trends(11,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Trends(12,:) - Trends(12,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch6')
ylim([min(Trends(11,:) - Trends(11,Mrk_Locs(1))) max(Trends(11,:) - Trends(11,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off

subplot(2,4,5)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Trends(13,:) - Trends(13,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Trends(14,:) - Trends(14,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch7')
ylim([min(Trends(13,:) - Trends(13,Mrk_Locs(1))) max(Trends(13,:) - Trends(13,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off

subplot(2,4,6)
plot(timestamps, Marker,'linewidth',1, 'Color', 'm');
hold on
plot(timestamps, Trends(15,:) - Trends(15,Mrk_Locs(1)),'linewidth',1, 'Color', 'R');
hold on
plot(timestamps, Trends(16,:) - Trends(16,Mrk_Locs(1)),'linewidth',1, 'Color', 'B');
title('Ch8')
ylim([min(Trends(15,:) - Trends(15,Mrk_Locs(1))) max(Trends(15,:) - Trends(15,Mrk_Locs(1)))]);
xlim([timestamps(Mrk_Locs(1,1)) timestamps(Mrk_Locs(1,6))])
hold off

% legend('HbO1', 'HbR1', 'HbO2', 'HbR2', 'HbO3', 'HbR3', 'HbO4', 'HbR4', ...
       % 'HbO5', 'HbR5', 'HbO6', 'HbR6', 'HbO7', 'HbR7', 'HbO8', 'HbR8', ...
       % 'Marker', 'Location', 'eastoutside');
end