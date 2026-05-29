function [Wavelet_Signals,Poly_Signals] = Fn_Filter(Raw_Signals, timestamps)
% function responsible for data filering and data preprocessing

%Wavelet order
    n=10; 
    w='db15'; % signal wavelet
    poly_degree = 3;

    % Empty matrix prep
    [num_channels, num_samples] = size(Raw_Signals);
    Poly_Signals = zeros(num_channels, num_samples);
    Wavelet_Signals = zeros(16, length(timestamps));

    disp(['Starting filters: Wavelet (db15, level 10) and polynomial ', num2str(poly_degree), '. order...']);

    for i = 1:num_channels
        signal = Raw_Signals(i, :);
        % 1. Wavelet transformation - filters pletysmo curve from data
        Wavelet_Signals(i,:) = wden(signal, 'rigrsure', 's', 'one', n, w);
        %2. Polynomial regression
        [p, ~, mu] = polyfit(timestamps, Wavelet_Signals(i, :), poly_degree);
        Poly_Signals(i, :) = polyval(p, timestamps, [], mu);
    end

    disp('Filter complete.');
end