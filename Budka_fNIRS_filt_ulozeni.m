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
figure();
plot(timestamps, HR, 'linewidth',1);

n=10; %wavelet order
w='db15'; % signal wavelet

figure();
signal1=Out(2,1000:end)-0.044;
trend1=wden(signal1,'rigrsure','s','one',n,w);
plot(timestamps, trend1,'linewidth',1);
hold on;

signal2=Out(3,1000:end)-0.02;;
trend2=wden(signal2,'rigrsure','s','one',n,w);
plot(timestamps, trend2,'linewidth',1);
hold on;

signal3=Out(4,1000:end)-0.047;
trend3=wden(signal3,'rigrsure','s','one',n,w);
plot(timestamps, trend3,'linewidth',1);
hold on;

signal4=Out(5,1000:end)-0.049;
trend4=wden(signal4,'rigrsure','s','one',n,w);
plot(timestamps, trend4,'linewidth',1);
hold on;

signal5=Out(6,1000:end)-0.049;
trend5=wden(signal5,'rigrsure','s','one',n,w);
plot(timestamps, trend5,'linewidth',1);
hold on;

signal6=Out(7,1000:end)-0.026;
trend6=wden(signal6,'rigrsure','s','one',n,w);
plot(timestamps, trend6,'linewidth',1);
hold on;

signal7=Out(8,1000:end)-0.048;
trend7=wden(signal7,'rigrsure','s','one',n,w);
plot(timestamps, trend7,'linewidth',1);
hold on;

signal8=Out(9,1000:end)-0.029;
trend8=wden(signal8,'rigrsure','s','one',n,w);
plot(timestamps, trend8,'linewidth',1);
hold on;

signal9=Out(10,1000:end)-0.045;
trend9=wden(signal9,'rigrsure','s','one',n,w);
plot(timestamps, trend9,'linewidth',1);
hold on;

signal10=Out(11,1000:end)-0.018;
trend10=wden(signal10,'rigrsure','s','one',n,w);
plot(timestamps, trend10,'linewidth',1);
hold on;

signal11=Out(12,1000:end)-0.034;
trend11=wden(signal11,'rigrsure','s','one',n,w);
plot(timestamps, trend11,'linewidth',1);
hold on;

signal12=Out(13,1000:end)-0.009;
trend12=wden(signal12,'rigrsure','s','one',n,w);
plot(timestamps, trend12,'linewidth',1);
hold on;

signal13=Out(14,1000:end)-0.037;
trend13=wden(signal13,'rigrsure','s','one',n,w);
plot(timestamps, trend13,'linewidth',1);
hold on;

signal14=Out(15,1000:end)-0.017;
trend14=wden(signal14,'rigrsure','s','one',n,w);
plot(timestamps, trend14,'linewidth',1);
hold on;

signal15=Out(16,1000:end)-0.047;
trend15=wden(signal15,'rigrsure','s','one',n,w);
plot(timestamps, trend15,'linewidth',1);
hold on;

signal16=Out(17,1000:end)-0.027;
trend16=wden(signal16,'rigrsure','s','one',n,w);
plot(timestamps, trend16,'linewidth',1);
hold on;

plot(timestamps, Marker,'linewidth',1);

legend('Trend 1', 'Trend 2', 'Trend 3', 'Trend 4', 'Trend 5', 'Trend 6', 'Trend 7', 'Trend 8', ...
       'Trend 9', 'Trend 10', 'Trend 11', 'Trend 12', 'Trend 13', 'Trend 14', 'Trend 15', 'Trend 16', ...
       'Marker', 'Location', 'eastoutside');

%% --- Export to excel ---
% Dialog box
answer = questdlg('Save data to excel?', 'Save selected', 'Yes','No','Yes');

if strcmp(answer, 'Yes')
    Mrk_Locs = find(Marker == 1);
    
    % File Name
    disp(['Saving data to file: ' full_path_to_excel]);
    
    % Data prep.
    len = length(timestamps);
    TrendsMatrix = zeros(16, len);
    fix_len = @(x, l) [x(1:min(length(x),l)), zeros(1, l-length(x))]; 
    
    TrendsMatrix(1,:) = fix_len(trend1, len);
    TrendsMatrix(2,:) = fix_len(trend2, len);
    TrendsMatrix(3,:) = fix_len(trend3, len);
    TrendsMatrix(4,:) = fix_len(trend4, len);
    TrendsMatrix(5,:) = fix_len(trend5, len);
    TrendsMatrix(6,:) = fix_len(trend6, len);
    TrendsMatrix(7,:) = fix_len(trend7, len);
    TrendsMatrix(8,:) = fix_len(trend8, len);
    TrendsMatrix(9,:) = fix_len(trend9, len);
    TrendsMatrix(10,:) = fix_len(trend10, len);
    TrendsMatrix(11,:) = fix_len(trend11, len);
    TrendsMatrix(12,:) = fix_len(trend12, len);
    TrendsMatrix(13,:) = fix_len(trend13, len);
    TrendsMatrix(14,:) = fix_len(trend14, len);
    TrendsMatrix(15,:) = fix_len(trend15, len);
    TrendsMatrix(16,:) = fix_len(trend16, len);
    
    RowNames = {'HR', 'Trend 1', 'Trend 2', 'Trend 3', 'Trend 4', 'Trend 5', 'Trend 6', 'Trend 7', 'Trend 8', ...
                'Trend 9', 'Trend 10', 'Trend 11', 'Trend 12', 'Trend 13', 'Trend 14', 'Trend 15', 'Trend 16'};
    PhaseNames = {'Relax (Start)', 'Cviceni Nohy', 'Relax (Mezi)', 'Cviceni Ruce', 'Relax (Konec)'};
    MarkerLabels = {'Start', 'Cviceni Nohy', 'Relax', 'Cviceni Ruce', 'Relax', 'Konec'};
    
    % Marker Table
    NewMarkerTable = cell(length(Mrk_Locs)+1, 4);
    NewMarkerTable(1,:) = {'Marker', 'Sample (Real)', 'Time (MM:SS)', 'Time(s)'};
    
    for i = 1:length(Mrk_Locs)
        if i <= length(MarkerLabels)
            lbl = MarkerLabels{i};
        else
            lbl = sprintf('Marker %d', i);
        end
        
        % Time calc. for MM:SS format
        time_val = timestamps(Mrk_Locs(i)); 
        mins = floor(time_val / 60);
        secs = round(time_val - mins*60);
        time_str = sprintf('00:%02d:%02d', mins, secs);
        
        NewMarkerTable{i+1, 1} = lbl;
        NewMarkerTable{i+1, 2} = Mrk_Locs(i) + 999; % Returns the true sample number, +999 offset because 1000 samples were cut from the start
        NewMarkerTable{i+1, 3} = time_str; % Returns the time MM:SS
        NewMarkerTable{i+1, 4} = mins*60 +secs; % Time in seconds as intiger
    end
    
    % Big data table
    No_Segments = length(Mrk_Locs) - 1; 
    BigHeader1 = cell(1, 1 + No_Segments*3); 
    BigHeader1{1} = 'Faze';
    BigHeader2 = cell(1, 1 + No_Segments*3);
    
    for p = 1:No_Segments
        col_idx = (p-1)*3 + 2;
        if p <= length(PhaseNames)
            BigHeader1{col_idx} = PhaseNames{p};
        else
            BigHeader1{col_idx} = sprintf('Phase %d', p);
        end
        BigHeader2{col_idx} = 'Max';
        BigHeader2{col_idx+1} = 'Min';
        BigHeader2{col_idx+2} = 'Mean';
    end
    
    DataBlock = cell(length(RowNames), 1 + No_Segments*3);
    DataBlock(:, 1) = RowNames';
    
    for p = 1:No_Segments
        idx_start = Mrk_Locs(p);
        idx_end = Mrk_Locs(p+1);
        col_idx = (p-1)*3 + 2;
        
        % HR
        seg_HR = HR(idx_start:min(idx_end, length(HR)));
        if ~isempty(seg_HR)
            % Používáme round(..., 2)
            DataBlock{1, col_idx} = max(seg_HR);
            DataBlock{1, col_idx+1} = min(seg_HR);
            DataBlock{1, col_idx+2} = mean(seg_HR);
        end
        % Trends
        for k = 1:16
            seg_Trend = TrendsMatrix(k, idx_start:min(idx_end, len));
            if ~isempty(seg_Trend)
                DataBlock{k+1, col_idx} = max(seg_Trend);
                DataBlock{k+1, col_idx+1} = min(seg_Trend);
                DataBlock{k+1, col_idx+2} = mean(seg_Trend);
            end
        end
    end
    
    % Write to excel
    try
        writecell(NewMarkerTable, full_path_to_excel, 'Sheet', 1, 'Range', 'F3');
        StartRow = 12; 
        writecell(BigHeader1, full_path_to_excel, 'Sheet', 1, 'Range', sprintf('B%d', StartRow));
        writecell(BigHeader2, full_path_to_excel, 'Sheet', 1, 'Range', sprintf('B%d', StartRow+1));
        writecell(DataBlock, full_path_to_excel, 'Sheet', 1, 'Range', sprintf('B%d', StartRow+2));
        
% Formatting Excel using (ActiveX)
        try
            Excel = actxserver('Excel.Application');
            Excel.Visible = 0;
            if exist(full_path_to_excel, 'file')
                Workbook = Excel.Workbooks.Open(full_path_to_excel);
            else
                 error('ERROR: Excel file not found!');
            end
            Sheet = Workbook.Sheets.Item(1);
            
            Sheet.Columns.AutoFit; 
            Sheet.Rows.AutoFit;
            
            % Check of cell dimensions if too small they are forced.
            cols_to_check = [2:26]; 
            for c = cols_to_check
                col = Sheet.Columns.Item(c);
                if col.ColumnWidth < 10
                    col.ColumnWidth = 10;
                end
            end
            
            % Centering text to center of cell (Big table)
            LastRow = Sheet.UsedRange.Rows.Count;
            RangeBigTable = Sheet.Range(sprintf('B12:Z%d', LastRow+2));
            RangeBigTable.HorizontalAlignment = -4108; % xlCenter = -4108
            RangeBigTable.NumberFormat = '0.00';
            
            % Centering text to center of cell (Small table)
            RangeSmallTable = Sheet.Range('F3:I10');
            RangeSmallTable.HorizontalAlignment = -4108;
            RangeSmallTable.NumberFormat = '0.00';
           
            % Bold titles
            Sheet.Range(sprintf('B12:Z13')).Font.Bold = 1;
            Sheet.Range('B3:I3').Font.Bold = 1;

            Workbook.Save;
            Workbook.Close;
            Excel.Quit;
            delete(Excel); 
            disp('Excel sucessfully formatted (Cell size + centering).');
        catch
            disp('ERROR: Automatic formatting failed (Missing Excel ActiveX).');
            disp('The data is saved but formatting must be done manually.');
        end
        
        msgbox('Data sucessfully saved and formatted', 'Finish');
    catch ME
        errordlg(['Data save failed! ' ME.message], 'ERROR');
    end
else
    disp('Data save canceled by user.');
end