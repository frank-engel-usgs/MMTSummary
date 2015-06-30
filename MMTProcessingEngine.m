function MMTProcessingEngine(inpath,infile,guiparams)
% Processes the MMT file and writes Excel spreadsheet
%
% Code by Dave S. Mueller
% Last modfied:
%    3/7/2013
%    Frank L. Engel, USGS, IL WSC

% Construct full file name
% ------------------------
filename = fullfile(inpath,filesep,infile);

% Initialize units conversion
% ---------------------------
if guiparams.metric_units
    qconvt  = 1;
    lvconvt = 1;
    aconvt  = 1;
else
    qconvt  = 35.31467;
    lvconvt = 3.28083;
    aconvt  = 10.76391;
end

% Convert the MMT to a MAT structure
% ----------------------------------
[...
    MMT,...
    MMT_Site_Info,...
    MMT_Transects,...
    MMT_Field_Config,...
    MMT_Active_Config,...
    MMT_Summary_None,...
    MMT_Summary_BT,...
    MMT_Summary_GGA,...
    MMT_Summary_VTG,...
    MMT_QAQC,...
    MMT_MB_Transects,...
    MMT_MB_Field_Config,...
    MMT_MB_Active_Config]       = mmt2mat(filename);

% Store data in mat file
% ----------------------
prefix   = [inpath infile(1:end-4)];
savefile = ['save (''' prefix ''')'];
if 1
    eval(savefile)
end

% Determine number of transects
% -----------------------------
%aa=length(MMT_Summary_None.BottomQ);
aa=sum(MMT_Summary_BT.Use);
idxTrans=find(MMT_Summary_BT.Use==1);
%number=1:aa;

% Create spreadsheet header row
% -----------------------------
head = createHeader;

% BT Reference
% ------------
if guiparams.bottom_track_reference
    data(1:aa,1)  = MMT_Summary_BT.FileName(idxTrans);
    if isnan(MMT_Site_Info.Name)
        data(1:aa,2) = {''};
    else
        data(1:aa,2) = cellstr(repmat(MMT_Site_Info.Name,aa,1));
    end
    if isnan(MMT_Site_Info.Number)
        data(1:aa,3) = {''};
    else
        data(1:aa,3) = cellstr(repmat(MMT_Site_Info.Number,aa,1));
    end
    if isnan(MMT_Site_Info.ADCPSerialNmb)
        data(1:aa,4) = {''};
    else
        data(1:aa,4) = cellstr(repmat(MMT_Site_Info.ADCPSerialNmb,aa,1));
    end
    if isnan(MMT_Active_Config.Wiz_Firmware)
        data(1:aa,5) = {''};
    else
        data(1:aa,5) = num2cell(MMT_Active_Config.Wiz_Firmware(idxTrans));
    end
    data(1:aa,6)    = num2cell(idxTrans);
    data(1:aa,7)    = cellstr(repmat('BT',aa,1));
    stimeconv       = MMT_Summary_BT.StartTime(idxTrans)./(60*60*24);
    stime           = datestr(719529+stimeconv);
    data(1:aa,8)    = cellstr(stime);
    etimeconv       = MMT_Summary_BT.EndTime(idxTrans)./(60*60*24);
    etime           = datestr(719529+etimeconv);
    data(1:aa,9)    = cellstr(etime);
    dursec          = MMT_Summary_BT.EndTime(idxTrans)-MMT_Summary_BT.StartTime(idxTrans);
    data(1:aa,10)   = num2cell(dursec);
    data(1:aa,11)   = num2cell(MMT_Summary_BT.TopQ(idxTrans).*qconvt);
    data(1:aa,12)   = num2cell(MMT_Summary_BT.MeasuredQ(idxTrans).*qconvt);
    data(1:aa,13)   = num2cell(MMT_Summary_BT.BottomQ(idxTrans).*qconvt);
    data(1:aa,14)   = num2cell(MMT_Summary_BT.LeftQ(idxTrans).*qconvt);
    data(1:aa,15)   = num2cell(MMT_Summary_BT.RightQ(idxTrans).*qconvt);
    data(1:aa,16)   = num2cell(MMT_Summary_BT.TotalQ(idxTrans).*qconvt);
    data(1:aa,17)   = num2cell(MMT_Summary_BT.Width(idxTrans).*lvconvt);
    data(1:aa,18)   = num2cell((MMT_Summary_BT.Width(idxTrans)-MMT_Summary_BT.LeftDistance(idxTrans)-MMT_Summary_BT.RightDistance(idxTrans)).*lvconvt);
    data(1:aa,19)   = num2cell(MMT_Summary_BT.TotalArea(idxTrans).*aconvt);
    if MMT_Summary_BT.TotalArea(idxTrans) > 0
        data(1:aa,20) = num2cell((MMT_Summary_BT.TotalQ(idxTrans)./MMT_Summary_BT.TotalArea(idxTrans)).*lvconvt);
    end
    data(1:aa,21)   = num2cell(MMT_Summary_BT.FlowDirection(idxTrans));
    data(1:aa,22)   = num2cell(MMT_Summary_BT.MaxWaterSpeed(idxTrans).*lvconvt);
    data(1:aa,23)   = num2cell(MMT_Summary_BT.MeanRiverVel(idxTrans).*lvconvt);
    data(1:aa,24)   = num2cell(MMT_Summary_BT.MaxWaterDepth(idxTrans).*lvconvt);
    data(1:aa,25)   = num2cell(MMT_Summary_BT.MeanWaterDepth(idxTrans).*lvconvt);
    data(1:aa,26)   = num2cell(MMT_Summary_BT.MeanBoatSpeed(idxTrans).*lvconvt);
    data(1:aa,27)   = num2cell(MMT_Summary_BT.MeanBoatCourse(idxTrans));
    data(1:aa,28)   = num2cell(MMT_Summary_BT.LeftDistance(idxTrans).*lvconvt);
    data(1:aa,29)   = num2cell(MMT_Summary_BT.RightDistance(idxTrans).*lvconvt);
    data(1:aa,30)   = num2cell(MMT_Summary_BT.LeftEdgeCoeff(idxTrans).*lvconvt);
    data(1:aa,31)   = num2cell(MMT_Summary_BT.RightEdgeCoeff(idxTrans));
    data(1:aa,32)   = num2cell(MMT_Summary_BT.TotalNmbEnsembles(idxTrans));
    data(1:aa,33)   = num2cell(MMT_Summary_BT.TotalBadEnsembles(idxTrans));
    data(1:aa,34)   = num2cell(MMT_Summary_BT.StartEnsemble(idxTrans));
    data(1:aa,35)   = num2cell(MMT_Summary_BT.EndEnsemble(idxTrans));
    data(1:aa,36)   = num2cell(MMT_Summary_BT.PercentGoodBins(idxTrans));
    data(1:aa,37)   = num2cell(MMT_Summary_BT.PowerCurveCoeff(idxTrans));
    if guiparams.metric_units
        data(1:aa,38) = num2cell(MMT_Summary_BT.ADCPTemperature(idxTrans));
    else
        data(1:aa,38) = num2cell(((9./5).*MMT_Summary_BT.ADCPTemperature(idxTrans))+32);
    end
    data(1:aa,39)   = num2cell(MMT_Summary_BT.BlankingDistance(idxTrans).*lvconvt);
    data(1:aa,40)   = num2cell(MMT_Summary_BT.BinSize(idxTrans).*lvconvt);
    data(1:aa,41)   = num2cell(MMT_Summary_BT.BTMode(idxTrans));
    data(1:aa,42)   = num2cell(MMT_Summary_BT.WTMode(idxTrans));
    data(1:aa,43)   = num2cell(MMT_Summary_BT.BTPings(idxTrans));
    data(1:aa,44)   = num2cell(MMT_Summary_BT.WTPings(idxTrans));
    data(1:aa,45)   = num2cell(MMT_Summary_BT.Begin_Left(idxTrans));
    data(1:aa,46)   = num2cell(MMT_Summary_BT.IsSubSectioned(idxTrans));
    data(1:aa,47)   = num2cell(MMT_Summary_BT.Use(idxTrans));
    data(1:aa,48)   = num2cell(MMT_Active_Config.Offsets_Transducer_Depth(idxTrans).*lvconvt);
    data(1:aa,49)   = num2cell(MMT_Active_Config.Proc_BT_Error_Vel_Threshold(idxTrans).*lvconvt);
    data(1:aa,50)   = num2cell(MMT_Active_Config.Proc_WT_Error_Vel_Threshold(idxTrans).*lvconvt);
    data(1:aa,51)   = num2cell(MMT_Active_Config.Proc_BT_Up_Vel_Threshold(idxTrans).*lvconvt);
    data(1:aa,52)   = num2cell(MMT_Active_Config.Proc_WT_Up_Vel_Threshold(idxTrans).*lvconvt);
    if isfield(MMT_Active_Config,'Wizard_Commands')
        temp        = MMT_Active_Config.Wizard_Commands(1,:);
        ncommands   = length(temp);
        wvidx       = 0;
        woidx       = 0;
        for j = 1:ncommands
            if cell2mat(strfind(temp(j),'WO')) == 1;
                woidx = j;
            end
            if cell2mat(strfind(temp(j),'WV')) == 1;
                wvidx = j;
            end
        end
        if wvidx > 0
            data(1:aa,53) = num2cell(repmat(str2double(strtrim(temp{wvidx}(3:end))),aa,1));
        else
            data(1:aa,53) = num2cell(zeros(aa,1));
        end
        if woidx > 0
            commaidx        = strfind(temp{woidx},',');
            data(1:aa,54)   = num2cell(repmat(str2double(temp{woidx}(3:commaidx-1)),aa,1));
            data(1:aa,55)   = num2cell(repmat(str2double(strtrim(temp{woidx}(commaidx+1:end))),aa,1));
        else
            data(1:aa,54)   = num2cell(zeros(aa,1));
            data(1:aa,55)   = num2cell(zeros(aa,1));
        end
    else
        data(1:aa,53)       ={''};
        data(1:aa,54)       ={''};
        data(1:aa,55)       ={''};
    end
    
    
    
    % Summary for database - dsm
    % --------------------------
    data2{1,1}      = 'Averages';
    data2{2,1}      = 'COV';
    data2{3,1}      = 'Totals';
    data2{3,6}      = sum(MMT_Summary_VTG.Use); % Number of Transects
    data2{3,8}      = data(1,8); % Start Time
    data2{3,9}      = data(data2{3,6},9); % End Time
    data2{3,10}     = sum([data{:,10}]); % Duration
    data2{1,11}     = mean([data{:,11}]); % Top Q
    data2{2,11}     = 100.*std([data{:,11}])./mean([data{:,11}]); % Top Q COV
    data2{1,12}     = mean([data{:,12}]); % Measured Q
    data2{2,12}     = 100.*std([data{:,12}])./mean([data{:,12}]); % Measured Q COV
    data2{1,13}     = mean([data{:,13}]); % Bottom Q
    data2{2,13}     = 100.*std([data{:,13}])./mean([data{:,13}]); % Bottom Q COV
    data2{1,14}     = mean([data{:,14}]); % Left Q
    data2{2,14}     = 100.*std([data{:,14}])./mean([data{:,14}]); % Left Q COV
    data2{1,15}     = mean([data{:,15}]); % Right Q
    data2{2,15}     = 100.*std([data{:,15}])./mean([data{:,15}]); % Right Q COV
    data2{1,16}     = mean([data{:,16}]); % Total Q
    data2{2,16}     = 100.*std([data{:,16}])./mean([data{:,16}]); % Total Q COV
    data2{1,17}     = mean([data{:,17}]); % Width
    data2{2,17}     = 100.*std([data{:,17}])./mean([data{:,17}]); % Width COV
    data2{1,19}     = mean([data{:,19}]); % Area
    data2{2,19}     = 100.*std([data{:,19}])./mean([data{:,19}]); % Area COV
    data2{1,25}     = mean([data{:,25}]); % Mean Depth
    data2{1,38}     = mean([data{:,38}]); % Temperature
    row             = sum(MMT_Summary_BT.Use)+3;
    row             = ['A',num2str(row)];
    
    
    % Write result to Excel
    % ---------------------
    savefilexls_bt=[prefix '.xlsx'];
    xlswrite (savefilexls_bt,head,  'BT', 'A1');
    xlswrite (savefilexls_bt,data,  'BT', 'A2');
    xlswrite (savefilexls_bt,data2, 'BT', row);
    
end

%
% GGA Reference
% -------------
if guiparams.gga_reference
    data(1:aa,1)  = MMT_Summary_GGA.FileName(idxTrans);
    if isnan(MMT_Site_Info.Name)
        data(1:aa,2) = {''};
    else
        data(1:aa,2) = cellstr(repmat(MMT_Site_Info.Name,aa,1));
    end
    if isnan(MMT_Site_Info.Number)
        data(1:aa,3) = {''};
    else
        data(1:aa,3) = cellstr(repmat(MMT_Site_Info.Number,aa,1));
    end
    if isnan(MMT_Site_Info.ADCPSerialNmb)
        data(1:aa,4) = {''};
    else
        data(1:aa,4) = cellstr(repmat(MMT_Site_Info.ADCPSerialNmb,aa,1));
    end
    if isnan(MMT_Active_Config.Wiz_Firmware)
        data(1:aa,5) = {''};
    else
        data(1:aa,5) = num2cell(MMT_Active_Config.Wiz_Firmware(idxTrans));
    end
    data(1:aa,6)    = num2cell(idxTrans);
    data(1:aa,7)    = cellstr(repmat('GGA',aa,1));
    stimeconv       = MMT_Summary_GGA.StartTime(idxTrans)./(60*60*24);
    stime           = datestr(719529+stimeconv,14);
    data(1:aa,8)    = cellstr(stime);
    etimeconv       = MMT_Summary_GGA.EndTime(idxTrans)./(60*60*24);
    etime           = datestr(719529+etimeconv,14);
    data(1:aa,9)    = cellstr(etime);
    dursec          = MMT_Summary_GGA.EndTime(idxTrans)-MMT_Summary_GGA.StartTime(idxTrans);
    data(1:aa,10)   = num2cell(dursec);
    data(1:aa,11)   = num2cell(MMT_Summary_GGA.TopQ(idxTrans).*qconvt);
    data(1:aa,12)   = num2cell(MMT_Summary_GGA.MeasuredQ(idxTrans).*qconvt);
    data(1:aa,13)   = num2cell(MMT_Summary_GGA.BottomQ(idxTrans).*qconvt);
    data(1:aa,14)   = num2cell(MMT_Summary_GGA.LeftQ(idxTrans).*qconvt);
    data(1:aa,15)   = num2cell(MMT_Summary_GGA.RightQ(idxTrans).*qconvt);
    data(1:aa,16)   = num2cell(MMT_Summary_GGA.TotalQ(idxTrans).*qconvt);
    data(1:aa,17)   = num2cell(MMT_Summary_GGA.Width(idxTrans).*lvconvt);
    data(1:aa,18)   = num2cell((MMT_Summary_GGA.Width(idxTrans)-MMT_Summary_GGA.LeftDistance(idxTrans)-MMT_Summary_GGA.RightDistance(idxTrans)).*lvconvt);
    data(1:aa,19)   = num2cell(MMT_Summary_GGA.TotalArea(idxTrans).*aconvt);
    if MMT_Summary_GGA.TotalArea(idxTrans) > 0
        data(1:aa,20) = num2cell((MMT_Summary_GGA.TotalQ(idxTrans)./MMT_Summary_GGA.TotalArea(idxTrans)).*lvconvt);
    end
    data(1:aa,21) = num2cell(MMT_Summary_GGA.FlowDirection(idxTrans));
    data(1:aa,22) = num2cell(MMT_Summary_GGA.MaxWaterSpeed(idxTrans).*lvconvt);
    data(1:aa,23) = num2cell(MMT_Summary_GGA.MeanRiverVel(idxTrans).*lvconvt);
    data(1:aa,24) = num2cell(MMT_Summary_GGA.MaxWaterDepth(idxTrans).*lvconvt);
    data(1:aa,25) = num2cell(MMT_Summary_GGA.MeanWaterDepth(idxTrans).*lvconvt);
    data(1:aa,26) = num2cell(MMT_Summary_GGA.MeanBoatSpeed(idxTrans).*lvconvt);
    data(1:aa,27) = num2cell(MMT_Summary_GGA.MeanBoatCourse(idxTrans));
    data(1:aa,28) = num2cell(MMT_Summary_GGA.LeftDistance(idxTrans).*lvconvt);
    data(1:aa,29) = num2cell(MMT_Summary_GGA.RightDistance(idxTrans).*lvconvt);
    data(1:aa,30) = num2cell(MMT_Summary_GGA.LeftEdgeCoeff(idxTrans).*lvconvt);
    data(1:aa,31) = num2cell(MMT_Summary_GGA.RightEdgeCoeff(idxTrans));
    data(1:aa,32) = num2cell(MMT_Summary_GGA.TotalNmbEnsembles(idxTrans));
    data(1:aa,33) = num2cell(MMT_Summary_GGA.TotalBadEnsembles(idxTrans));
    data(1:aa,34) = num2cell(MMT_Summary_GGA.StartEnsemble(idxTrans));
    data(1:aa,35) = num2cell(MMT_Summary_GGA.EndEnsemble(idxTrans));
    data(1:aa,36) = num2cell(MMT_Summary_GGA.PercentGoodBins(idxTrans));
    data(1:aa,37) = num2cell(MMT_Summary_GGA.PowerCurveCoeff(idxTrans));
    if guiparams.metric_units
        data(1:aa,38) = num2cell(MMT_Summary_GGA.ADCPTemperature(idxTrans));
    else
        data(1:aa,38) = num2cell(((9./5).*MMT_Summary_GGA.ADCPTemperature(idxTrans))+32);
    end
    data(1:aa,39) = num2cell(MMT_Summary_GGA.BlankingDistance(idxTrans).*lvconvt);
    data(1:aa,40) = num2cell(MMT_Summary_GGA.BinSize(idxTrans).*lvconvt);
    data(1:aa,41) = num2cell(MMT_Summary_GGA.BTMode(idxTrans));
    data(1:aa,42) = num2cell(MMT_Summary_GGA.WTMode(idxTrans));
    data(1:aa,43) = num2cell(MMT_Summary_GGA.BTPings(idxTrans));
    data(1:aa,44) = num2cell(MMT_Summary_GGA.WTPings(idxTrans));
    data(1:aa,45) = num2cell(MMT_Summary_GGA.Begin_Left(idxTrans));
    data(1:aa,46) = num2cell(MMT_Summary_GGA.IsSubSectioned(idxTrans));
    data(1:aa,47) = num2cell(MMT_Summary_GGA.Use(idxTrans));
    data(1:aa,48) = num2cell(MMT_Active_Config.Offsets_Transducer_Depth(idxTrans).*lvconvt);
    data(1:aa,49) = num2cell(MMT_Active_Config.Proc_BT_Error_Vel_Threshold(idxTrans).*lvconvt);
    data(1:aa,50) = num2cell(MMT_Active_Config.Proc_WT_Error_Vel_Threshold(idxTrans).*lvconvt);
    data(1:aa,51) = num2cell(MMT_Active_Config.Proc_BT_Up_Vel_Threshold(idxTrans).*lvconvt);
    data(1:aa,52) = num2cell(MMT_Active_Config.Proc_WT_Up_Vel_Threshold(idxTrans).*lvconvt);
    if isfield(MMT_Active_Config,'Wizard_Commands')
        temp        = MMT_Active_Config.Wizard_Commands(1,:);
        ncommands   = length(temp);
        wvidx       = 0;
        woidx       = 0;
        for j = 1:ncommands
            if cell2mat(strfind(temp(j),'WO'))==1;
                woidx = j;
            end
            if cell2mat(strfind(temp(j),'WV'))==1;
                wvidx = j;
            end
        end
        if wvidx > 0
            data(1:aa,53) = num2cell(repmat(str2double(strtrim(temp{wvidx}(3:end))),aa,1));
        else
            data(1:aa,53) = num2cell(zeros(aa,1));
        end
        if woidx>0
            commaidx      = strfind(temp{woidx},',');
            data(1:aa,54) = num2cell(repmat(str2double(temp{woidx}(3:commaidx-1)),aa,1));
            data(1:aa,55) = num2cell(repmat(str2double(strtrim(temp{woidx}(commaidx+1:end))),aa,1));
        else
            data(1:aa,54) = num2cell(zeros(aa,1));
            data(1:aa,55) = num2cell(zeros(aa,1));
        end
    else
        data(1:aa,53) = {''};
        data(1:aa,54) = {''};
        data(1:aa,55) = {''};
    end
    
    
    % Summary for database - dsm
    % --------------------------
    data2{1,1}  ='Averages';
    data2{2,1}  ='COV';
    data2{3,1}  ='Totals';
    data2{3,6}  = sum(MMT_Summary_VTG.Use); % Number of Transects
    data2{3,8}  = data(1,8); % Start Time
    data2{3,9}  = data(data2{3,6},9); % End Time
    data2{3,10} = sum([data{:,10}]); % Duration
    data2{1,11} = mean([data{:,11}]); % Top Q
    data2{2,11} = 100.*std([data{:,11}])./mean([data{:,11}]); % Top Q COV
    data2{1,12} = mean([data{:,12}]); % Measured Q
    data2{2,12} = 100.*std([data{:,12}])./mean([data{:,12}]); % Measured Q COV
    data2{1,13} = mean([data{:,13}]); % Bottom Q
    data2{2,13} = 100.*std([data{:,13}])./mean([data{:,13}]); % Bottom Q COV
    data2{1,14} = mean([data{:,14}]); % Left Q
    data2{2,14} = 100.*std([data{:,14}])./mean([data{:,14}]); % Left Q COV
    data2{1,15} = mean([data{:,15}]); % Right Q
    data2{2,15} = 100.*std([data{:,15}])./mean([data{:,15}]); % Right Q COV
    data2{1,16} = mean([data{:,16}]); % Total Q
    data2{2,16} = 100.*std([data{:,16}])./mean([data{:,16}]); % Total Q COV
    data2{1,17} = mean([data{:,17}]); % Width
    data2{2,17} = 100.*std([data{:,17}])./mean([data{:,17}]); % Width COV
    data2{1,19} = mean([data{:,19}]); % Area
    data2{2,19} = 100.*std([data{:,19}])./mean([data{:,19}]); % Area COV
    data2{1,25} = mean([data{:,25}]); % Mean Depth
    data2{1,38} = mean([data{:,38}]); % Temperature
    row         = sum(MMT_Summary_GGA.Use)+3;
    row         = ['A',num2str(row)];
    
    
    % Write Excel spreadsheet
    % -----------------------
    savefilexls_gga = [prefix '.xlsx'];
    xlswrite (savefilexls_gga,head, 'GGA', 'A1');
    xlswrite (savefilexls_gga,data, 'GGA', 'A2');
    xlswrite (savefilexls_gga,data2, 'GGA', row);
end

% VTG Reference
% -------------
if guiparams.vtg_reference
    data(1:aa,1)  = MMT_Summary_VTG.FileName(idxTrans);
    if isnan(MMT_Site_Info.Name)
        data(1:aa,2) = {''};
    else
        data(1:aa,2) = cellstr(repmat(MMT_Site_Info.Name,aa,1));
    end
    if isnan(MMT_Site_Info.Number)
        data(1:aa,3) = {''};
    else
        data(1:aa,3) = cellstr(repmat(MMT_Site_Info.Number,aa,1));
    end
    if isnan(MMT_Site_Info.ADCPSerialNmb)
        data(1:aa,4) = {''};
    else
        data(1:aa,4) = cellstr(repmat(MMT_Site_Info.ADCPSerialNmb,aa,1));
    end
    if isnan(MMT_Active_Config.Wiz_Firmware)
        data(1:aa,5) = {''};
    else
        data(1:aa,5) = num2cell(MMT_Active_Config.Wiz_Firmware(idxTrans));
    end
    data(1:aa,6)  = num2cell(idxTrans);
    data(1:aa,7)  = cellstr(repmat('VTG',aa,1));
    stimeconv     = MMT_Summary_VTG.StartTime(idxTrans)./(60*60*24);
    stime         = datestr(719529+stimeconv,14);
    data(1:aa,8)  = cellstr(stime);
    etimeconv     = MMT_Summary_VTG.EndTime(idxTrans)./(60*60*24);
    etime         = datestr(719529+etimeconv,14);
    data(1:aa,9)  = cellstr(etime);
    dursec        = MMT_Summary_VTG.EndTime(idxTrans)-MMT_Summary_VTG.StartTime(idxTrans);
    data(1:aa,10) = num2cell(dursec);
    data(1:aa,11) = num2cell(MMT_Summary_VTG.TopQ(idxTrans).*qconvt);
    data(1:aa,12) = num2cell(MMT_Summary_VTG.MeasuredQ(idxTrans).*qconvt);
    data(1:aa,13) = num2cell(MMT_Summary_VTG.BottomQ(idxTrans).*qconvt);
    data(1:aa,14) = num2cell(MMT_Summary_VTG.LeftQ(idxTrans).*qconvt);
    data(1:aa,15) = num2cell(MMT_Summary_VTG.RightQ(idxTrans).*qconvt);
    data(1:aa,16) = num2cell(MMT_Summary_VTG.TotalQ(idxTrans).*qconvt);
    data(1:aa,17) = num2cell(MMT_Summary_VTG.Width(idxTrans).*lvconvt);
    data(1:aa,18) = num2cell((MMT_Summary_VTG.Width(idxTrans)-MMT_Summary_VTG.LeftDistance(idxTrans)-MMT_Summary_VTG.RightDistance(idxTrans)).*lvconvt);
    data(1:aa,19) = num2cell(MMT_Summary_VTG.TotalArea(idxTrans).*aconvt);
    if MMT_Summary_VTG.TotalArea(idxTrans) > 0
        data(1:aa,20) = num2cell((MMT_Summary_VTG.TotalQ(idxTrans)./MMT_Summary_VTG.TotalArea(idxTrans)).*lvconvt);
    end
    data(1:aa,21) = num2cell(MMT_Summary_VTG.FlowDirection(idxTrans));
    data(1:aa,22) = num2cell(MMT_Summary_VTG.MaxWaterSpeed(idxTrans).*lvconvt);
    data(1:aa,23) = num2cell(MMT_Summary_VTG.MeanRiverVel(idxTrans).*lvconvt);
    data(1:aa,24) = num2cell(MMT_Summary_VTG.MaxWaterDepth(idxTrans).*lvconvt);
    data(1:aa,25) = num2cell(MMT_Summary_VTG.MeanWaterDepth(idxTrans).*lvconvt);
    data(1:aa,26) = num2cell(MMT_Summary_VTG.MeanBoatSpeed(idxTrans).*lvconvt);
    data(1:aa,27) = num2cell(MMT_Summary_VTG.MeanBoatCourse(idxTrans));
    data(1:aa,28) = num2cell(MMT_Summary_VTG.LeftDistance(idxTrans).*lvconvt);
    data(1:aa,29) = num2cell(MMT_Summary_VTG.RightDistance(idxTrans).*lvconvt);
    data(1:aa,30) = num2cell(MMT_Summary_VTG.LeftEdgeCoeff(idxTrans).*lvconvt);
    data(1:aa,31) = num2cell(MMT_Summary_VTG.RightEdgeCoeff(idxTrans));
    data(1:aa,32) = num2cell(MMT_Summary_VTG.TotalNmbEnsembles(idxTrans));
    data(1:aa,33) = num2cell(MMT_Summary_VTG.TotalBadEnsembles(idxTrans));
    data(1:aa,34) = num2cell(MMT_Summary_VTG.StartEnsemble(idxTrans));
    data(1:aa,35) = num2cell(MMT_Summary_VTG.EndEnsemble(idxTrans));
    data(1:aa,36) = num2cell(MMT_Summary_VTG.PercentGoodBins(idxTrans));
    data(1:aa,37) = num2cell(MMT_Summary_VTG.PowerCurveCoeff(idxTrans));
    if guiparams.metric_units
        data(1:aa,38) = num2cell(MMT_Summary_VTG.ADCPTemperature(idxTrans));
    else
        data(1:aa,38) = num2cell(((9./5).*MMT_Summary_VTG.ADCPTemperature(idxTrans))+32);
    end
    data(1:aa,39) = num2cell(MMT_Summary_VTG.BlankingDistance(idxTrans).*lvconvt);
    data(1:aa,40) = num2cell(MMT_Summary_VTG.BinSize(idxTrans).*lvconvt);
    data(1:aa,41) = num2cell(MMT_Summary_VTG.BTMode(idxTrans));
    data(1:aa,42) = num2cell(MMT_Summary_VTG.WTMode(idxTrans));
    data(1:aa,43) = num2cell(MMT_Summary_VTG.BTPings(idxTrans));
    data(1:aa,44) = num2cell(MMT_Summary_VTG.WTPings(idxTrans));
    data(1:aa,45) = num2cell(MMT_Summary_VTG.Begin_Left(idxTrans));
    data(1:aa,46) = num2cell(MMT_Summary_VTG.IsSubSectioned(idxTrans));
    data(1:aa,47) = num2cell(MMT_Summary_VTG.Use(idxTrans));
    data(1:aa,48) = num2cell(MMT_Active_Config.Offsets_Transducer_Depth(idxTrans).*lvconvt);
    data(1:aa,49) = num2cell(MMT_Active_Config.Proc_BT_Error_Vel_Threshold(idxTrans).*lvconvt);
    data(1:aa,50) = num2cell(MMT_Active_Config.Proc_WT_Error_Vel_Threshold(idxTrans).*lvconvt);
    data(1:aa,51) = num2cell(MMT_Active_Config.Proc_BT_Up_Vel_Threshold(idxTrans).*lvconvt);
    data(1:aa,52) = num2cell(MMT_Active_Config.Proc_WT_Up_Vel_Threshold(idxTrans).*lvconvt);
    if isfield(MMT_Active_Config,'Wizard_Commands')
        temp      = MMT_Active_Config.Wizard_Commands(1,:);
        ncommands = length(temp);
        wvidx     = 0;
        woidx     = 0;
        for j = 1:ncommands
            if cell2mat(strfind(temp(j),'WO'))==1;
                woidx = j;
            end
            if cell2mat(strfind(temp(j),'WV'))==1;
                wvidx = j;
            end
        end
        if wvidx > 0
            data(1:aa,53) = num2cell(repmat(str2double(strtrim(temp{wvidx}(3:end))),aa,1));
        else
            data(1:aa,53) = num2cell(zeros(aa,1));
        end
        if woidx > 0
            commaidx      = strfind(temp{woidx},',');
            data(1:aa,54) = num2cell(repmat(str2double(temp{woidx}(3:commaidx-1)),aa,1));
            data(1:aa,55) = num2cell(repmat(str2double(strtrim(temp{woidx}(commaidx+1:end))),aa,1));
        else
            data(1:aa,54) = num2cell(zeros(aa,1));
            data(1:aa,55) = num2cell(zeros(aa,1));
        end
    else
        data(1:aa,53) = {''};
        data(1:aa,54) = {''};
        data(1:aa,55) = {''};
    end
    savefilexls_vtg = [prefix '.xlsx'];
    xlswrite (savefilexls_vtg,head, 'VTG', 'A1');
    xlswrite (savefilexls_vtg,data, 'VTG', 'A2');
    
    % Summary for database - dsm
    % --------------------------
    data2{1,1}  = 'Averages';
    data2{2,1}  = 'COV';
    data2{3,1}  = 'Totals';
    data2{3,6}  = sum(MMT_Summary_VTG.Use); % Number of Transects
    data2{3,8}  = data(1,8); % Start Time
    data2{3,9}  = data(data2{3,6},9); % End Time
    data2{3,10} = sum([data{:,10}]); % Duration
    data2{1,11} = mean([data{:,11}]); % Top Q
    data2{2,11} = 100.*std([data{:,11}])./mean([data{:,11}]); % Top Q COV
    data2{1,12} = mean([data{:,12}]); % Measured Q
    data2{2,12} = 100.*std([data{:,12}])./mean([data{:,12}]); % Measured Q COV
    data2{1,13} = mean([data{:,13}]); % Bottom Q
    data2{2,13} = 100.*std([data{:,13}])./mean([data{:,13}]); % Bottom Q COV
    data2{1,14} = mean([data{:,14}]); % Left Q
    data2{2,14} = 100.*std([data{:,14}])./mean([data{:,14}]); % Left Q COV
    data2{1,15} = mean([data{:,15}]); % Right Q
    data2{2,15} = 100.*std([data{:,15}])./mean([data{:,15}]); % Right Q COV
    data2{1,16} = mean([data{:,16}]); % Total Q
    data2{2,16} = 100.*std([data{:,16}])./mean([data{:,16}]); % Total Q COV
    data2{1,17} = mean([data{:,17}]); % Width
    data2{2,17} = 100.*std([data{:,17}])./mean([data{:,17}]); % Width COV
    data2{1,19} = mean([data{:,19}]); % Area
    data2{2,19} = 100.*std([data{:,19}])./mean([data{:,19}]); % Area COV
    data2{1,25} = mean([data{:,25}]); % Mean Depth
    data2{1,38} = mean([data{:,38}]); % Temperature
    row         = sum(MMT_Summary_VTG.Use)+3;
    row         = ['A',num2str(row)];
    
    % Write Excel spreadsheet
    % -----------------------
    savefilexls_vtg = [prefix '.xlsx'];
    xlswrite (savefilexls_vtg,head, 'VTG', 'A1');
    xlswrite (savefilexls_vtg,data, 'VTG', 'A2');
    xlswrite (savefilexls_vtg,data2, 'VTG', row);
    
end

%%%%%%%%%%%%%%%%
% SUBFUNCTIONS %
%%%%%%%%%%%%%%%%
function head = createHeader

% Initialize spreadsheet headers
% ------------------------------
head{1,1}  = 'File Name';
head{1,2}  = 'Location';
head{1,3}  = 'Site ID';
head{1,4}  = 'SN';
head{1,5}  = 'Firmware';
head{1,6}  = 'Number';
head{1,7}  = 'Reference';
head{1,8}  = 'Start Date/Time';
head{1,9}  = 'End Date/Time';
head{1,10}  = 'Duration [s]';
head{1,11} = 'Top Q';
head{1,12} = 'Meas Q';
head{1,13} = 'Bot Q';
head{1,14} = 'Left Q';
head{1,15} = 'Right Q';
head{1,16} = 'Total Q';
head{1,17} = 'Total Width';
head{1,18} = 'Meas Width';
head{1,19} = 'Total Area';
head{1,20} = 'Mean Velocity';
head{1,21} = 'Flow Direction';
head{1,22} = 'Max Water Speed';
head{1,23} = 'Mean River Velocity';
head{1,24} = 'Max Water Depth';
head{1,25} = 'Mean Water Depth';
head{1,26} = 'Mean Boat Speed';
head{1,27} = 'Mean Boat Course';
head{1,28} = 'Left Distance';
head{1,29} = 'Right Distance';
head{1,30} = 'Left Edge Slope Coeff';
head{1,31} = 'Right Edge Slope Coeff';
head{1,32} = 'Total Number of Ensembles';
head{1,33} = 'Total Bad Ensembles';
head{1,34} = 'Start Ensemble';
head{1,35} = 'End Emsemble';
head{1,36} = 'Percent Good Bins';
head{1,37} = 'Power Curve Coeff';
head{1,38} = 'ADCP Temperature';
head{1,39} = 'Blanking Distance';
head{1,40} = 'Bin Size';
head{1,41} = 'BT Mode';
head{1,42} = 'WT Mode';
head{1,43} = 'BT Pings';
head{1,44} = 'WT Pings';
head{1,45} = 'Begin Left';
head{1,46} = 'Is Sub Sectioned';
head{1,47} = 'Use in Summary';
head{1,48} = 'ADCP Transducer Depth';
head{1,49} = 'BT Error Velocity Threshhold';
head{1,50} = 'WT Error Velocity Threshhold';
head{1,51} = 'BT Up Velocity Threshhold';
head{1,52} = 'WT Up Velocity Threshhold';
head{1,53} = 'WV';
head{1,54} = 'WO Subpings';
head{1,55} = 'WO Time Between Subpings';
% [EOF] createHeader
