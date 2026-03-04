function [] = Fn_save2excel(Trends, Marker, timestamps, full_path_to_excel)
%This function shows a dialog box asking the user to save data to an excel
%file. If user selects yes, max, min, mean and variance and markers data
%will be saved to an excel file with the same name as the loaded file

%% Dialog box
answer = questdlg('Save data to excel?', 'Save selected', 'Yes','No','Yes');

if strcmp(answer, 'Yes')
    Mrk_Locs = find(Marker == 1);
    
    % File Name
    disp(['Saving data to file: ' full_path_to_excel]);
    
    % Data prep.
    len = length(timestamps);
    TrendsMatrix = zeros(16, len);
    fix_len = @(x, l) [x(1:min(length(x),l)), zeros(1, l-length(x))]; 
    
    TrendsMatrix(1,:) = fix_len(Trends(1), len);
    TrendsMatrix(2,:) = fix_len(Trends(2), len);
    TrendsMatrix(3,:) = fix_len(Trends(3), len);
    TrendsMatrix(4,:) = fix_len(Trends(4), len);
    TrendsMatrix(5,:) = fix_len(Trends(5), len);
    TrendsMatrix(6,:) = fix_len(Trends(6), len);
    TrendsMatrix(7,:) = fix_len(Trends(7), len);
    TrendsMatrix(8,:) = fix_len(Trends(8), len);
    TrendsMatrix(9,:) = fix_len(Trends(9), len);
    TrendsMatrix(10,:) = fix_len(Trends(10), len);
    TrendsMatrix(11,:) = fix_len(Trends(11), len);
    TrendsMatrix(12,:) = fix_len(Trends(12), len);
    TrendsMatrix(13,:) = fix_len(Trends(13), len);
    TrendsMatrix(14,:) = fix_len(Trends(14), len);
    TrendsMatrix(15,:) = fix_len(Trends(15), len);
    TrendsMatrix(16,:) = fix_len(Trends(16), len);
    
    RowNames = {'HR', 'HbO1', 'HbR1', 'HbO2', 'HbR2', 'HbO3', 'HbR3', 'HbO4', 'HbR4', 'HbO5', 'HbR5', 'HbO6', 'HbR6', 'HbO7', 'HbR7',... 
        'HbO8', 'HbR8', 'HbO9', 'HbR9', 'HbO10', 'HbR10', 'HbO11', 'HbR11', 'HbO12', 'HbR12', 'HbO13', 'HbR13', 'HbO14', 'HbR14', 'HbO15', 'HbR15', 'HbO16', 'HbR16',};
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
    BigHeader1 = cell(1, 1 + No_Segments*4); 
    BigHeader1{1} = 'Faze';
    BigHeader2 = cell(1, 1 + No_Segments*4);
    
    for p = 1:No_Segments
        col_idx = (p-1)*4 + 2;
        if p <= length(PhaseNames)
            BigHeader1{col_idx} = PhaseNames{p};
        else
            BigHeader1{col_idx} = sprintf('Phase %d', p);
        end
        BigHeader2{col_idx} = 'Max';
        BigHeader2{col_idx+1} = 'Min';
        BigHeader2{col_idx+2} = 'Mean';
        BigHeader2{col_idx+3} = 'Variance';
    end
    
    DataBlock = cell(length(RowNames), 1 + No_Segments*4);
    DataBlock(:, 1) = RowNames';
    
    for p = 1:No_Segments
        idx_start = Mrk_Locs(p);
        idx_end = Mrk_Locs(p+1);
        col_idx = (p-1)*4 + 2;
        
        % HR
        seg_HR = HR(idx_start:min(idx_end, length(HR)));
        if ~isempty(seg_HR)
            DataBlock{1, col_idx} = max(seg_HR);
            DataBlock{1, col_idx+1} = min(seg_HR);
            DataBlock{1, col_idx+2} = mean(seg_HR);
            DataBlock{1, col_idx+3} = var(seg_HR);
        end
        % Trends
        for k = 1:16
            seg_Trend = TrendsMatrix(k, idx_start:min(idx_end, len));
            if ~isempty(seg_Trend)
                DataBlock{k+1, col_idx} = max(seg_Trend);
                DataBlock{k+1, col_idx+1} = min(seg_Trend);
                DataBlock{k+1, col_idx+2} = mean(seg_Trend);
                DataBlock{k+1, col_idx+3} = var(seg_Trend);
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
end