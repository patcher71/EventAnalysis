classdef YWaveAnalyzerAppv2 < matlab.apps.AppBase

    % Properties that correspond to app components
    properties (Access = public)
        UIFigure                  matlab.ui.Figure
        GridLayout                matlab.ui.container.GridLayout
        LeftPanel                 matlab.ui.container.Panel
        SelectRootDirButton       matlab.ui.control.Button
        RootDirLabel              matlab.ui.control.Label
        CellFoldersLabel          matlab.ui.control.Label
        CellFolderDropDown        matlab.ui.control.DropDown
        YWaveFoldersLabel         matlab.ui.control.Label
        YWaveFolderListBox        matlab.ui.control.ListBox
        AnalyzeSelectedButton     matlab.ui.control.Button
        AveragingPanel            matlab.ui.container.Panel
        AnalyzedWavesLabel        matlab.ui.control.Label
        AnalyzedWavesListBox      matlab.ui.control.ListBox
        CreateAverageCDFButton    matlab.ui.control.Button
        ClearAveragesButton       matlab.ui.control.Button
        ExportToExcelButton       matlab.ui.control.Button
        SummaryTable              matlab.ui.control.Table
        RightPanel                matlab.ui.container.Panel
        TabGroup                  matlab.ui.container.TabGroup
        IndividualTab             matlab.ui.container.Tab
        PeakAxes                  matlab.ui.control.UIAxes
        IEIAxes                   matlab.ui.control.UIAxes
        AveragedTab               matlab.ui.container.Tab
        AvgPeakAxes               matlab.ui.control.UIAxes
        AvgIEIAxes                matlab.ui.control.UIAxes
    end

    properties (Access = private)
        RootDirectory = '';
        CurrentCellFolder = '';
        AvailableFolders = {};
        PeakCDFData = {};  % Store CDF data for export
        IEICDFData = {};   % Store CDF data for export
        AnalyzedWaveNames = {};  % Store names of analyzed waves
        AveragedPeakData = {};  % Store averaged peak CDF data
        AveragedIEIData = {};   % Store averaged IEI CDF data
        AverageGroupCount = 0;  % Counter for averaged groups
    end

    methods (Access = private)

        % Button pushed function: SelectRootDirButton
        function SelectRootDirButtonPushed(app, event)
            % Let user select root directory (e.g., sIPSC Analysis folder)
            selectedDir = uigetdir('', 'Select Root Directory (e.g., sIPSC Analysis)');
            if selectedDir == 0
                return;
            end
            
            app.RootDirectory = selectedDir;
            app.RootDirLabel.Text = ['Root: ' selectedDir];
            
            % Find all Cell folders
            dirContents = dir(app.RootDirectory);
            folderNames = {dirContents([dirContents.isdir]).name};
            % Filter for Cell folders (starts with "Cell")
            cellFolders = folderNames(startsWith(folderNames, 'Cell'));
            
            if isempty(cellFolders)
                uialert(app.UIFigure, 'No Cell folders found in selected directory.', 'Warning');
                app.CellFolderDropDown.Items = {};
                app.YWaveFolderListBox.Items = {};
                return;
            end
            
            % Update dropdown with cell folders
            app.CellFolderDropDown.Items = cellFolders;
            app.CellFolderDropDown.Value = cellFolders{1};
            
            % Load YWave folders for first cell
            CellFolderDropDownValueChanged(app, event);
        end

        % Value changed function: CellFolderDropDown
        function CellFolderDropDownValueChanged(app, event)
            selectedCell = app.CellFolderDropDown.Value;
            
            if isempty(selectedCell)
                return;
            end
            
            % Construct path to eventer.output folder
            eventerPath = fullfile(app.RootDirectory, selectedCell, 'eventer.output');
            
            if ~isfolder(eventerPath)
                uialert(app.UIFigure, sprintf('eventer.output folder not found in %s', selectedCell), 'Warning');
                app.YWaveFolderListBox.Items = {};
                return;
            end
            
            app.CurrentCellFolder = eventerPath;
            
            % Find all YWave folders in eventer.output
            dirContents = dir(eventerPath);
            folderNames = {dirContents([dirContents.isdir]).name};
            yWaveFolders = folderNames(startsWith(folderNames, 'Data_ch1_YWave'));
            
            app.AvailableFolders = yWaveFolders;
            app.YWaveFolderListBox.Items = yWaveFolders;
            
            if isempty(yWaveFolders)
                uialert(app.UIFigure, sprintf('No YWave folders found in %s/eventer.output', selectedCell), 'Warning');
            end
        end

        % Button pushed function: AnalyzeSelectedButton
        function AnalyzeSelectedButtonPushed(app, event)
            selectedFolders = app.YWaveFolderListBox.Value;
            
            if isempty(selectedFolders)
                uialert(app.UIFigure, 'Please select at least one YWave folder.', 'Warning');
                return;
            end
            
            % Clear previous plots and data
            cla(app.PeakAxes);
            cla(app.IEIAxes);
            hold(app.PeakAxes, 'on');
            hold(app.IEIAxes, 'on');
            
            % Prepare summary data
            numFolders = length(selectedFolders);
            summaryData = cell(numFolders, 4);
            colors = lines(numFolders);
            
            % Initialize storage for CDF data
            app.PeakCDFData = cell(numFolders, 3); % Wave name, values, probabilities
            app.IEICDFData = cell(numFolders, 3);
            app.AnalyzedWaveNames = cell(numFolders, 1);
            
            % Process each selected folder
            for i = 1:numFolders
                folderName = selectedFolders{i};
                folderPath = fullfile(app.CurrentCellFolder, folderName);
                
                % Read and parse summary.txt
                summaryPath = fullfile(folderPath, 'summary.txt');
                [totalEvents, meanAmp, eventFreq] = parseSummary(summaryPath);
                
                % Store in table
                waveLabel = sprintf('Wave%d', i);
                summaryData{i, 1} = waveLabel;
                summaryData{i, 2} = totalEvents;
                summaryData{i, 3} = meanAmp;
                summaryData{i, 4} = eventFreq;
                
                app.AnalyzedWaveNames{i} = waveLabel;
                
                % Read peak.txt and convert to pA
                peakPath = fullfile(folderPath, 'TXT', 'peak.txt');
                peakData = readmatrix(peakPath);
                peakData_pA = peakData * 1e12; % Convert A to pA
                
                % Read IEI.txt
                ieiPath = fullfile(folderPath, 'TXT', 'IEI.txt');
                ieiData = readmatrix(ieiPath);
                ieiData = ieiData(~isnan(ieiData)); % Remove NaN values
                
                % Calculate and store CDF data
                [sortedPeak, peakProb] = calculateCDF(peakData_pA);
                [sortedIEI, ieiProb] = calculateCDF(ieiData);
                
                app.PeakCDFData{i, 1} = sprintf('Wave%d', i);
                app.PeakCDFData{i, 2} = sortedPeak;
                app.PeakCDFData{i, 3} = peakProb;
                
                app.IEICDFData{i, 1} = sprintf('Wave%d', i);
                app.IEICDFData{i, 2} = sortedIEI;
                app.IEICDFData{i, 3} = ieiProb;
                
                % Plot cumulative distributions
                plot(app.PeakAxes, sortedPeak, peakProb, 'DisplayName', sprintf('Wave%d', i), ...
                     'Color', colors(i,:), 'LineWidth', 1.5);
                plot(app.IEIAxes, sortedIEI, ieiProb, 'DisplayName', sprintf('Wave%d', i), ...
                     'Color', colors(i,:), 'LineWidth', 1.5);
            end
            
            % Update summary table
            app.SummaryTable.Data = summaryData;
            
            % Format axes
            xlabel(app.PeakAxes, 'Peak Amplitude (pA)');
            ylabel(app.PeakAxes, 'Cumulative Probability');
            title(app.PeakAxes, 'Cumulative Distribution of Peak Amplitudes');
            legend(app.PeakAxes, 'Location', 'best');
            grid(app.PeakAxes, 'on');
            
            xlabel(app.IEIAxes, 'Inter-Event Interval (s)');
            ylabel(app.IEIAxes, 'Cumulative Probability');
            title(app.IEIAxes, 'Cumulative Distribution of IEIs');
            legend(app.IEIAxes, 'Location', 'best');
            grid(app.IEIAxes, 'on');
            
            hold(app.PeakAxes, 'off');
            hold(app.IEIAxes, 'off');
            
            % Enable export button and update analyzed waves listbox
            app.ExportToExcelButton.Enable = 'on';
            app.AnalyzedWavesListBox.Items = app.AnalyzedWaveNames;
            app.CreateAverageCDFButton.Enable = 'on';
        end

        % Button pushed function: CreateAverageCDFButton
        function CreateAverageCDFButtonPushed(app, event)
            selectedWaves = app.AnalyzedWavesListBox.Value;
            
            if isempty(selectedWaves)
                uialert(app.UIFigure, 'Please select waves to average.', 'Warning');
                return;
            end
            
            if length(selectedWaves) < 2
                uialert(app.UIFigure, 'Please select at least 2 waves to average.', 'Warning');
                return;
            end
            
            % Get indices of selected waves
            waveIndices = [];
            waveNumbers = [];
            for i = 1:length(selectedWaves)
                idx = find(strcmp(app.AnalyzedWaveNames, selectedWaves{i}));
                if ~isempty(idx)
                    waveIndices(end+1) = idx;
                    % Extract wave number from name (e.g., "Wave3" -> 3)
                    waveNum = str2double(regexp(selectedWaves{i}, '\d+', 'match'));
                    waveNumbers(end+1) = waveNum;
                end
            end
            
            % Create label showing which waves were averaged
            if length(waveNumbers) <= 5
                % If 5 or fewer waves, list them all
                waveStr = sprintf('%d,', waveNumbers);
                waveStr = waveStr(1:end-1); % Remove trailing comma
                groupLabel = sprintf('Waves %s (n=%d)', waveStr, length(waveNumbers));
            else
                % If more than 5 waves, show range
                groupLabel = sprintf('Waves %d-%d (n=%d)', min(waveNumbers), max(waveNumbers), length(waveNumbers));
            end
            
            % Create averaged CDFs using interpolation
            [avgPeakX, avgPeakY, sdPeakY] = averageCDFs(app.PeakCDFData(waveIndices, :));
            [avgIEIX, avgIEIY, sdIEIY] = averageCDFs(app.IEICDFData(waveIndices, :));
            
            % Increment group counter
            app.AverageGroupCount = app.AverageGroupCount + 1;
            
            % Store averaged data
            newIdx = size(app.AveragedPeakData, 1) + 1;
            app.AveragedPeakData{newIdx, 1} = groupLabel;
            app.AveragedPeakData{newIdx, 2} = avgPeakX;
            app.AveragedPeakData{newIdx, 3} = avgPeakY;
            app.AveragedPeakData{newIdx, 4} = sdPeakY;
            
            app.AveragedIEIData{newIdx, 1} = groupLabel;
            app.AveragedIEIData{newIdx, 2} = avgIEIX;
            app.AveragedIEIData{newIdx, 3} = avgIEIY;
            app.AveragedIEIData{newIdx, 4} = sdIEIY;
            
            % Plot on averaged axes
            hold(app.AvgPeakAxes, 'on');
            hold(app.AvgIEIAxes, 'on');
            
            colors = lines(newIdx);
            
            % Plot mean with shaded error bars
            plotCDFWithError(app.AvgPeakAxes, avgPeakX, avgPeakY, sdPeakY, groupLabel, colors(newIdx,:));
            plotCDFWithError(app.AvgIEIAxes, avgIEIX, avgIEIY, sdIEIY, groupLabel, colors(newIdx,:));
            
            % Format averaged axes
            xlabel(app.AvgPeakAxes, 'Peak Amplitude (pA)');
            ylabel(app.AvgPeakAxes, 'Cumulative Probability');
            title(app.AvgPeakAxes, 'Averaged Cumulative Distribution of Peak Amplitudes');
            legend(app.AvgPeakAxes, 'Location', 'best');
            grid(app.AvgPeakAxes, 'on');
            
            xlabel(app.AvgIEIAxes, 'Inter-Event Interval (s)');
            ylabel(app.AvgIEIAxes, 'Cumulative Probability');
            title(app.AvgIEIAxes, 'Averaged Cumulative Distribution of IEIs');
            legend(app.AvgIEIAxes, 'Location', 'best');
            grid(app.AvgIEIAxes, 'on');
            
            hold(app.AvgPeakAxes, 'off');
            hold(app.AvgIEIAxes, 'off');
            
            app.ClearAveragesButton.Enable = 'on';
        end

        % Button pushed function: ClearAveragesButton
        function ClearAveragesButtonPushed(app, event)
            % Clear averaged data and plots
            app.AveragedPeakData = {};
            app.AveragedIEIData = {};
            app.AverageGroupCount = 0;
            
            % Clear all children from axes (including shaded regions)
            cla(app.AvgPeakAxes, 'reset');
            cla(app.AvgIEIAxes, 'reset');
            
            app.ClearAveragesButton.Enable = 'off';
        end

        % Button pushed function: ExportToExcelButton
        function ExportToExcelButtonPushed(app, event)
            if isempty(app.SummaryTable.Data)
                uialert(app.UIFigure, 'No data to export. Please analyze folders first.', 'Warning');
                return;
            end
            
            % Get filename from user - default to root directory
            defaultFilename = fullfile(app.RootDirectory, 'YWave_Analysis.xlsx');
            [file, path] = uiputfile('*.xlsx', 'Save Excel File', defaultFilename);
            if file == 0
                return; % User cancelled
            end
            
            fullFilePath = fullfile(path, file);
            
            % Create waitbar
            wb = uiprogressdlg(app.UIFigure, 'Title', 'Exporting to Excel', ...
                'Message', 'Writing summary data...', 'Indeterminate', 'on');
            
            try
                % Sheet 1: Summary Statistics
                summaryHeaders = {'Folder', 'Total Events', 'Mean Amplitude (pA)', 'Frequency (Hz)'};
                summaryTable = [summaryHeaders; app.SummaryTable.Data];
                writecell(summaryTable, fullFilePath, 'Sheet', 'Summary Statistics');
                
                % Sheet 2: Peak Amplitude CDF
                wb.Message = 'Writing peak amplitude data...';
                peakSheet = createCDFSheet(app.PeakCDFData, 'Peak Amplitude (pA)');
                writecell(peakSheet, fullFilePath, 'Sheet', 'Peak Amplitude CDF');
                
                % Sheet 3: IEI CDF
                wb.Message = 'Writing IEI data...';
                ieiSheet = createCDFSheet(app.IEICDFData, 'IEI (s)');
                writecell(ieiSheet, fullFilePath, 'Sheet', 'IEI CDF');
                
                % Sheet 4 & 5: Averaged CDFs (if any)
                if ~isempty(app.AveragedPeakData)
                    wb.Message = 'Writing averaged peak data...';
                    avgPeakSheet = createAveragedCDFSheet(app.AveragedPeakData, 'Peak Amplitude (pA)');
                    writecell(avgPeakSheet, fullFilePath, 'Sheet', 'Averaged Peak CDF');
                    
                    wb.Message = 'Writing averaged IEI data...';
                    avgIEISheet = createAveragedCDFSheet(app.AveragedIEIData, 'IEI (s)');
                    writecell(avgIEISheet, fullFilePath, 'Sheet', 'Averaged IEI CDF');
                end
                
                close(wb);
                uialert(app.UIFigure, sprintf('Data exported successfully to:\n%s', fullFilePath), ...
                    'Export Complete', 'Icon', 'success');
                
            catch ME
                close(wb);
                uialert(app.UIFigure, sprintf('Error exporting data:\n%s', ME.message), ...
                    'Export Failed', 'Icon', 'error');
            end
        end
    end

    % Component initialization
    methods (Access = private)

        % Create UIFigure and components
        function createComponents(app)

            % Create UIFigure and hide until all components are created
            app.UIFigure = uifigure('Visible', 'off');
            app.UIFigure.Position = [100 100 1200 700];
            app.UIFigure.Name = 'YWave Data Analyzer';

            % Create GridLayout
            app.GridLayout = uigridlayout(app.UIFigure);
            app.GridLayout.ColumnWidth = {'1x', '2x'};
            app.GridLayout.RowHeight = {'1x'};

            % Create LeftPanel
            app.LeftPanel = uipanel(app.GridLayout);
            app.LeftPanel.Title = 'Controls';
            app.LeftPanel.Layout.Row = 1;
            app.LeftPanel.Layout.Column = 1;

            % Create SelectRootDirButton
            app.SelectRootDirButton = uibutton(app.LeftPanel, 'push');
            app.SelectRootDirButton.ButtonPushedFcn = createCallbackFcn(app, @SelectRootDirButtonPushed, true);
            app.SelectRootDirButton.Position = [10 650 330 30];
            app.SelectRootDirButton.Text = 'Select Root Directory';

            % Create RootDirLabel
            app.RootDirLabel = uilabel(app.LeftPanel);
            app.RootDirLabel.Position = [10 615 330 30];
            app.RootDirLabel.Text = 'Root: Not selected';
            app.RootDirLabel.FontSize = 10;

            % Create CellFoldersLabel
            app.CellFoldersLabel = uilabel(app.LeftPanel);
            app.CellFoldersLabel.FontWeight = 'bold';
            app.CellFoldersLabel.Position = [10 585 330 22];
            app.CellFoldersLabel.Text = 'Select Cell Folder:';

            % Create CellFolderDropDown
            app.CellFolderDropDown = uidropdown(app.LeftPanel);
            app.CellFolderDropDown.Items = {'No cells available'};
            app.CellFolderDropDown.ValueChangedFcn = createCallbackFcn(app, @CellFolderDropDownValueChanged, true);
            app.CellFolderDropDown.Position = [10 555 330 30];

            % Create YWaveFoldersLabel
            app.YWaveFoldersLabel = uilabel(app.LeftPanel);
            app.YWaveFoldersLabel.FontWeight = 'bold';
            app.YWaveFoldersLabel.Position = [10 525 330 22];
            app.YWaveFoldersLabel.Text = 'Available YWave Folders:';

            % Create YWaveFolderListBox
            app.YWaveFolderListBox = uilistbox(app.LeftPanel);
            app.YWaveFolderListBox.Multiselect = 'on';
            app.YWaveFolderListBox.Position = [10 420 330 100];

            % Create AnalyzeSelectedButton
            app.AnalyzeSelectedButton = uibutton(app.LeftPanel, 'push');
            app.AnalyzeSelectedButton.ButtonPushedFcn = createCallbackFcn(app, @AnalyzeSelectedButtonPushed, true);
            app.AnalyzeSelectedButton.Position = [10 390 330 30];
            app.AnalyzeSelectedButton.Text = 'Analyze Selected Folders';
            app.AnalyzeSelectedButton.FontWeight = 'bold';

            % Create AveragingPanel
            app.AveragingPanel = uipanel(app.LeftPanel);
            app.AveragingPanel.Title = 'Create Averaged CDFs';
            app.AveragingPanel.Position = [10 210 330 170];

            % Create AnalyzedWavesLabel
            app.AnalyzedWavesLabel = uilabel(app.AveragingPanel);
            app.AnalyzedWavesLabel.Text = 'Select waves to average:';
            app.AnalyzedWavesLabel.Position = [10 120 300 22];

            % Create AnalyzedWavesListBox
            app.AnalyzedWavesListBox = uilistbox(app.AveragingPanel);
            app.AnalyzedWavesListBox.Multiselect = 'on';
            app.AnalyzedWavesListBox.Items = {};
            app.AnalyzedWavesListBox.Position = [10 50 310 70];

            % Create CreateAverageCDFButton
            app.CreateAverageCDFButton = uibutton(app.AveragingPanel, 'push');
            app.CreateAverageCDFButton.ButtonPushedFcn = createCallbackFcn(app, @CreateAverageCDFButtonPushed, true);
            app.CreateAverageCDFButton.Position = [10 15 150 30];
            app.CreateAverageCDFButton.Text = 'Create Average';
            app.CreateAverageCDFButton.Enable = 'off';

            % Create ClearAveragesButton
            app.ClearAveragesButton = uibutton(app.AveragingPanel, 'push');
            app.ClearAveragesButton.ButtonPushedFcn = createCallbackFcn(app, @ClearAveragesButtonPushed, true);
            app.ClearAveragesButton.Position = [170 15 150 30];
            app.ClearAveragesButton.Text = 'Clear Averages';
            app.ClearAveragesButton.Enable = 'off';

            % Create ExportToExcelButton
            app.ExportToExcelButton = uibutton(app.LeftPanel, 'push');
            app.ExportToExcelButton.ButtonPushedFcn = createCallbackFcn(app, @ExportToExcelButtonPushed, true);
            app.ExportToExcelButton.Position = [10 170 330 30];
            app.ExportToExcelButton.Text = 'Export to Excel';
            app.ExportToExcelButton.BackgroundColor = [0.2 0.6 0.2];
            app.ExportToExcelButton.FontColor = [1 1 1];
            app.ExportToExcelButton.Enable = 'off';

            % Create SummaryTable
            app.SummaryTable = uitable(app.LeftPanel);
            app.SummaryTable.ColumnName = {'Folder'; 'Total Events'; 'Mean Amp (pA)'; 'Freq (Hz)'};
            app.SummaryTable.RowName = {};
            app.SummaryTable.Position = [10 10 330 150];
            app.SummaryTable.ColumnWidth = {60, 90, 90, 70};

            % Create RightPanel
            app.RightPanel = uipanel(app.GridLayout);
            app.RightPanel.Title = 'Cumulative Distributions';
            app.RightPanel.Layout.Row = 1;
            app.RightPanel.Layout.Column = 2;

            % Create TabGroup
            app.TabGroup = uitabgroup(app.RightPanel);
            app.TabGroup.Position = [10 10 760 660];

            % Create IndividualTab
            app.IndividualTab = uitab(app.TabGroup);
            app.IndividualTab.Title = 'Individual CDFs';

            % Create PeakAxes
            app.PeakAxes = uiaxes(app.IndividualTab);
            app.PeakAxes.Position = [20 330 720 290];

            % Create IEIAxes
            app.IEIAxes = uiaxes(app.IndividualTab);
            app.IEIAxes.Position = [20 10 720 290];

            % Create AveragedTab
            app.AveragedTab = uitab(app.TabGroup);
            app.AveragedTab.Title = 'Averaged CDFs';

            % Create AvgPeakAxes
            app.AvgPeakAxes = uiaxes(app.AveragedTab);
            app.AvgPeakAxes.Position = [20 330 720 290];

            % Create AvgIEIAxes
            app.AvgIEIAxes = uiaxes(app.AveragedTab);
            app.AvgIEIAxes.Position = [20 10 720 290];

            % Show the figure after all components are created
            app.UIFigure.Visible = 'on';
        end
    end

    % App creation and deletion
    methods (Access = public)

        % Construct app
        function app = YWaveAnalyzerAppv2

            % Create UIFigure and components
            createComponents(app)

            % Register the app with App Designer
            registerApp(app, app.UIFigure)

            if nargout == 0
                clear app
            end
        end

        % Code that executes before app deletion
        function delete(app)

            % Delete UIFigure when app is deleted
            delete(app.UIFigure)
        end
    end
end

% Helper functions
function [totalEvents, meanAmp, eventFreq] = parseSummary(filepath)
    % Parse summary.txt file to extract key parameters
    fid = fopen(filepath, 'r');
    if fid == -1
        error('Cannot open file: %s', filepath);
    end
    
    totalEvents = NaN;
    meanAmp = NaN;
    eventFreq = NaN;
    
    while ~feof(fid)
        line = fgetl(fid);
        
        % Extract total events
        if contains(line, 'Total number of events detected:')
            tokens = regexp(line, 'detected:\s*(\d+)', 'tokens');
            if ~isempty(tokens)
                totalEvents = str2double(tokens{1}{1});
            end
        end
        
        % Extract mean amplitude (already in pA)
        if contains(line, 'Mean event amplitude (pA):')
            tokens = regexp(line, '\(pA\):\s*([\d.]+)', 'tokens');
            if ~isempty(tokens)
                meanAmp = str2double(tokens{1}{1});
            end
        end
        
        % Extract event frequency
        if contains(line, 'Event frequency (in Hz):')
            tokens = regexp(line, 'Hz\):\s*([\d.]+)', 'tokens');
            if ~isempty(tokens)
                eventFreq = str2double(tokens{1}{1});
            end
        end
    end
    
    fclose(fid);
end

function [sortedData, cumulativeProb] = calculateCDF(data)
    % Calculate cumulative distribution function
    sortedData = sort(data);
    n = length(sortedData);
    cumulativeProb = (1:n)' / n;
end

function sheetData = createCDFSheet(cdfData, valueLabel)
    % Create a cell array for Excel sheet with CDF data
    % Each wave gets two columns: values and cumulative probability
    
    numWaves = size(cdfData, 1);
    
    % Find maximum length needed
    maxLen = 0;
    for i = 1:numWaves
        maxLen = max(maxLen, length(cdfData{i, 2}));
    end
    
    % Create headers
    headers = cell(1, numWaves * 2);
    for i = 1:numWaves
        headers{2*i-1} = sprintf('%s %s', cdfData{i,1}, valueLabel);
        headers{2*i} = sprintf('%s Cum Prob', cdfData{i,1});
    end
    
    % Create data matrix
    dataMatrix = cell(maxLen, numWaves * 2);
    for i = 1:numWaves
        values = cdfData{i, 2};
        probs = cdfData{i, 3};
        dataLen = length(values);
        
        % Fill in the data
        for j = 1:dataLen
            dataMatrix{j, 2*i-1} = values(j);
            dataMatrix{j, 2*i} = probs(j);
        end
        
        % Fill remaining cells with empty
        for j = (dataLen+1):maxLen
            dataMatrix{j, 2*i-1} = '';
            dataMatrix{j, 2*i} = '';
        end
    end
    
    % Combine headers and data
    sheetData = [headers; dataMatrix];
end

function [avgX, avgY, sdY] = averageCDFs(cdfData)
    % Average multiple CDFs using interpolation to common X values
    numCurves = size(cdfData, 1);
    
    % Find overall min and max X values across all curves
    allX = [];
    for i = 1:numCurves
        allX = [allX; cdfData{i, 2}(:)];
    end
    minX = min(allX);
    maxX = max(allX);
    
    % Create common X values for interpolation (500 points)
    avgX = linspace(minX, maxX, 500)';
    
    % Interpolate each CDF to common X values
    interpY = zeros(length(avgX), numCurves);
    for i = 1:numCurves
        xData = cdfData{i, 2};
        yData = cdfData{i, 3};
        
        % Remove duplicate X values by keeping unique values
        [xUnique, uniqueIdx] = unique(xData, 'stable');
        yUnique = yData(uniqueIdx);
        
        % Handle edge case where all X values are the same
        if length(xUnique) < 2
            % Just use constant interpolation
            interpY(:, i) = yUnique(1);
        else
            % Interpolate, using nearest neighbor extrapolation
            interpY(:, i) = interp1(xUnique, yUnique, avgX, 'linear', 'extrap');
            
            % Clamp to [0, 1] range (in case of extrapolation issues)
            interpY(:, i) = max(0, min(1, interpY(:, i)));
        end
    end
    
    % Calculate mean and SD across curves
    avgY = mean(interpY, 2);
    sdY = std(interpY, 0, 2);
end

function sheetData = createAveragedCDFSheet(avgData, valueLabel)
    % Create Excel sheet for averaged CDF data with mean and SD
    numGroups = size(avgData, 1);
    
    % Find maximum length needed
    maxLen = 0;
    for i = 1:numGroups
        maxLen = max(maxLen, length(avgData{i, 2}));
    end
    
    % Create headers
    headers = cell(1, numGroups * 3);
    for i = 1:numGroups
        headers{3*i-2} = sprintf('%s %s', avgData{i,1}, valueLabel);
        headers{3*i-1} = sprintf('%s Mean Prob', avgData{i,1});
        headers{3*i} = sprintf('%s SD Prob', avgData{i,1});
    end
    
    % Create data matrix
    dataMatrix = cell(maxLen, numGroups * 3);
    for i = 1:numGroups
        xVals = avgData{i, 2};
        meanVals = avgData{i, 3};
        sdVals = avgData{i, 4};
        dataLen = length(xVals);
        
        % Fill in the data
        for j = 1:dataLen
            dataMatrix{j, 3*i-2} = xVals(j);
            dataMatrix{j, 3*i-1} = meanVals(j);
            dataMatrix{j, 3*i} = sdVals(j);
        end
        
        % Fill remaining cells with empty
        for j = (dataLen+1):maxLen
            dataMatrix{j, 3*i-2} = '';
            dataMatrix{j, 3*i-1} = '';
            dataMatrix{j, 3*i} = '';
        end
    end
    
    % Combine headers and data
    sheetData = [headers; dataMatrix];
end

function plotCDFWithError(ax, xData, yMean, ySD, labelStr, color)
    % Plot CDF with shaded error region (mean ± SD)
    
    % Create filled area for error bars
    xFill = [xData; flipud(xData)];
    yFill = [yMean + ySD; flipud(yMean - ySD)];
    
    % Remove any NaN or out-of-range values
    yFill = max(0, min(1, yFill));
    
    fill(ax, xFill, yFill, color, 'FaceAlpha', 0.2, 'EdgeColor', 'none', 'HandleVisibility', 'off');
    
    % Plot mean line
    plot(ax, xData, yMean, 'Color', color, 'LineWidth', 2, 'DisplayName', labelStr);
end

function plotCDF(ax, data, labelStr, color)
    % Plot cumulative distribution function
    sortedData = sort(data);
    n = length(sortedData);
    cumulativeProb = (1:n)' / n;
    
    plot(ax, sortedData, cumulativeProb, 'DisplayName', labelStr, ...
         'Color', color, 'LineWidth', 1.5);
end
