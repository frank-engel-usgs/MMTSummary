function varargout = Parse_MMT_GUI(varargin)
% PARSE_MMT_GUI M-file for Parse_MMT_GUI.fig
%      PARSE_MMT_GUI, by itself, creates a new PARSE_MMT_GUI or raises the existing
%      singleton*.
%
%      H = PARSE_MMT_GUI returns the handle to a new PARSE_MMT_GUI or the handle to
%      the existing singleton*.
%
%      PARSE_MMT_GUI('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in PARSE_MMT_GUI.M with the given input arguments.
%
%      PARSE_MMT_GUI('Property','Value',...) creates a new PARSE_MMT_GUI or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before Parse_MMT_GUI_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to Parse_MMT_GUI_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help Parse_MMT_GUI

% Last Modified by GUIDE v2.5 02-Feb-2011 11:10:41

% Edited by J. Boldt, 2/13/2012

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Parse_MMT_GUI_OpeningFcn, ...
                   'gui_OutputFcn',  @Parse_MMT_GUI_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before Parse_MMT_GUI is made visible.
function Parse_MMT_GUI_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to Parse_MMT_GUI (see VARARGIN)

% Choose default command line output for Parse_MMT_GUI
handles.output = hObject;
%set(handles.Mode_pannel,'SelectionChangeFcn',@pushbutton1_Callback);
handles.reprocess=0;
% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Parse_MMT_GUI wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Parse_MMT_GUI_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
 handles = guidata(hObject); 

% 
% Start error handler
% -------------------
try
%
% Set path to current directory
% ------------------------------
mmtpath = pwd;
%
% If previous path has been saved use that path
% ---------------------------------------------
if exist('LastDir.mat') == 2
    load('LastDir.mat');
    %
    % Fix for old saved files that may have mmtpath=0
    % -----------------------------------------------
    if mmtpath==0
        mmtpath=pwd;
    end
end
%
% Save path in case user cancels from uigetfile
% ---------------------------------------------
mmtpathsave=mmtpath;
%
% Get file
% --------
if handles.reprocess==0
    [file,mmtpath] = uigetfile({'*.mmt','All .mmt Files'; '*.*','All Files'},'Select .mmt File',mmtpath);
    handles.file=file;
    handles.mmtpath=mmtpath;
else
    file=handles.file;
    mmtpath=handles.mmtpath;
end
% 
% If use cancelled reset path and don't execute mmt2mat
% -----------------------------------------------------
if mmtpath==0
    mmtpath=mmtpathsave;
else
    % 
    % If a file was selected generate full filename and store path
    % ------------------------------------------------------------
    infile = [mmtpath file];    
    if exist('LastDir.mat') == 2
        save('LastDir.mat','mmtpath','-append');
    else
        save('LastDir.mat','mmtpath');
    end
    %
    % Disable button while processing
    % -------------------------------
    set(handles.pushbutton1,'Enable','off')
    drawnow;
    %
    % Process file
    % ------------
    if handles.reprocess==0
        set(handles.status,'String','Reading MMT File')
        drawnow;
        [MMT,MMT_Site_Info,MMT_Transects,MMT_Field_Config,MMT_Active_Config,MMT_Summary_None,MMT_Summary_BT,MMT_Summary_GGA,MMT_Summary_VTG,MMT_QAQC,MMT_MB_Transects,MMT_MB_Field_Config,MMT_MB_Active_Config]=mmt2mat(infile);
        handles.MMT_Site_Info=MMT_Site_Info;
        handles.MMT_Active_Config=MMT_Active_Config;
        handles.MMT_Summary_None=MMT_Summary_None;
        handles.MMT_Summary_BT=MMT_Summary_BT;
        handles.MMT_Summary_GGA=MMT_Summary_GGA;
        handles.MMT_Summary_VTG=MMT_Summary_VTG;
    else
        handles.reprocess=0;
        MMT_Site_Info=handles.MMT_Site_Info;
        MMT_Active_Config=handles.MMT_Active_Config;
        MMT_Summary_None=handles.MMT_Summary_None;
        MMT_Summary_BT=handles.MMT_Summary_BT;
        MMT_Summary_GGA=handles.MMT_Summary_GGA;
        MMT_Summary_VTG=handles.MMT_Summary_VTG;
    end
    set(handles.status,'String','Creating Excel File')
    drawnow;
    %
    % Store data in mat file
    % ----------------------
    prefix=[mmtpath file(1:end-4)];
    savefile=['save (''' prefix ''')'];
    eval(savefile)
    %
    % Determine number of transects
    % -----------------------------
    %aa=length(MMT_Summary_None.BottomQ);
    aa=sum(MMT_Summary_BT.Use);
    idxTrans=find(MMT_Summary_BT.Use==1);
    %number=1:aa;
    %
    % Initialize units conversion
    % ---------------------------
    if get(handles.unitsmetric,'Value')
        qconvt=1;
        lvconvt=1;
        aconvt=1;
    else
        qconvt=35.31467;
        lvconvt=3.28083;
        aconvt=10.76391;
    end    
    %
    % Initialize spreadsheet headers
    % ------------------------------
    head{1,1} = 'File Name';
    head{1,2}  = 'Location';
    head{1,3}  = 'Site ID';
    head{1,4}  = 'SN';
    head{1,5}  = 'Number';
    head{1,6}  = 'Reference';
    head{1,7}  = 'Start Time';
    head{1,8}  = 'End Time';
    head{1,9}  = 'Duration [s]';
    head{1,10} = 'Top Q';
    head{1,11} = 'Mid Q';
    head{1,12} = 'Bot Q';
    head{1,13} = 'Left Q';
    head{1,14} = 'Right Q';
    head{1,15} = 'Total Q';
    head{1,16} = 'Width';
    head{1,17} = 'Total Area';
    head{1,18} = 'Mean Velocity';
    head{1,19} = 'Flow Direction';
    head{1,20} = 'Max Water Speed';
    head{1,21} = 'Mean River Velocity';
    head{1,22} = 'Max Water Depth';
    head{1,23} = 'Mean Water Depth';
    head{1,24} = 'Mean Boat Speed';
    head{1,25} = 'Mean Boat Course';
    head{1,26} = 'Left Distance';
    head{1,27} = 'Right Distance';
    head{1,28} = 'Left Edge Slope Coeff';
    head{1,29} = 'Right Edge Slope Coeff';
    head{1,30} = 'Total Number of Ensembles';
    head{1,31} = 'Total Bad Ensembles';
    head{1,32} = 'Start Ensemble';
    head{1,33} = 'End Emsemble';
    head{1,34} = 'Percent Good Bins';
    head{1,35} = 'Power Curve Coeff';
    head{1,36} = 'ADCP Temperature';
    head{1,37} = 'Blanking Distance';
    head{1,38} = 'Bin Size';
    head{1,39} = 'BT Mode';
    head{1,40} = 'WT Mode';
    head{1,41} = 'BT Pings';
    head{1,42} = 'WT Pings';
    head{1,43} = 'Begin Left';
    head{1,44} = 'Is Sub Sectioned';
    head{1,45} = 'Use in Summary';
    head{1,46} = 'ADCP Transducer Depth';
    head{1,47} = 'BT Error Velocity Threshhold';
    head{1,48} = 'WT Error Velocity Threshhold';
    head{1,49} = 'BT Up Velocity Threshhold';
    head{1,50} = 'WT Up Velocity Threshhold';
    head{1,51} = 'WV';
    head{1,52} = 'WO Subpings';
    head{1,53} = 'WO Time Between Subpings';
    %
    % Create spreadsheet for selected reference
    % =========================================
    %
    % BT Reference
    % ------------
    if get(handles.BTradio,'Value')
        data(1:aa,1)  = MMT_Summary_BT.FileName(idxTrans);
        if isnan(MMT_Site_Info.Name)
            data(1:aa,2)={''};
        else
            data(1:aa,2) = cellstr(repmat(MMT_Site_Info.Name,aa,1));
        end
        if isnan(MMT_Site_Info.Number)
            data(1:aa,3)={''};
        else
            data(1:aa,3) = cellstr(repmat(MMT_Site_Info.Number,aa,1));
        end
        if isnan(MMT_Site_Info.ADCPSerialNmb)
            data(1:aa,4)={''};
        else
            data(1:aa,4) = cellstr(repmat(MMT_Site_Info.ADCPSerialNmb,aa,1));
        end
        data(1:aa,5)  = num2cell(idxTrans);
        data(1:aa,6)  = cellstr(repmat('BT',aa,1));
        stimeconv=MMT_Summary_BT.StartTime(idxTrans)./(60*60*24);
        stime=datestr(719529+stimeconv,14);
        data(1:aa,7)  = cellstr(stime);
        etimeconv=MMT_Summary_BT.EndTime(idxTrans)./(60*60*24);
        etime=datestr(719529+etimeconv,14);
        data(1:aa,8)  = cellstr(etime);
        dursec=MMT_Summary_BT.EndTime(idxTrans)-MMT_Summary_BT.StartTime(idxTrans);
        data(1:aa,9)  = num2cell(dursec);
        data(1:aa,10) = num2cell(MMT_Summary_BT.TopQ(idxTrans).*qconvt);
        data(1:aa,11) = num2cell(MMT_Summary_BT.MeasuredQ(idxTrans).*qconvt);
        data(1:aa,12) = num2cell(MMT_Summary_BT.BottomQ(idxTrans).*qconvt);
        data(1:aa,13) = num2cell(MMT_Summary_BT.LeftQ(idxTrans).*qconvt);
        data(1:aa,14) = num2cell(MMT_Summary_BT.RightQ(idxTrans).*qconvt);
        data(1:aa,15) = num2cell(MMT_Summary_BT.TotalQ(idxTrans).*qconvt);
        data(1:aa,16) = num2cell(MMT_Summary_BT.Width(idxTrans).*lvconvt);
        data(1:aa,17) = num2cell(MMT_Summary_BT.TotalArea(idxTrans).*aconvt);
        if MMT_Summary_BT.TotalArea(idxTrans)>0
            data(1:aa,18) =num2cell((MMT_Summary_BT.TotalQ(idxTrans)./MMT_Summary_BT.TotalArea(idxTrans)).*lvconvt);
        end
        data(1:aa,19) = num2cell(MMT_Summary_BT.FlowDirection(idxTrans));
        data(1:aa,20) = num2cell(MMT_Summary_BT.MaxWaterSpeed(idxTrans).*lvconvt);
        data(1:aa,21) = num2cell(MMT_Summary_BT.MeanRiverVel(idxTrans).*lvconvt);
        data(1:aa,22) = num2cell(MMT_Summary_BT.MaxWaterDepth(idxTrans).*lvconvt);
        data(1:aa,23) = num2cell(MMT_Summary_BT.MeanWaterDepth(idxTrans).*lvconvt);
        data(1:aa,24) = num2cell(MMT_Summary_BT.MeanBoatSpeed(idxTrans).*lvconvt);
        data(1:aa,25) = num2cell(MMT_Summary_BT.MeanBoatCourse(idxTrans));
        data(1:aa,26) = num2cell(MMT_Summary_BT.LeftDistance(idxTrans).*lvconvt);
        data(1:aa,27) = num2cell(MMT_Summary_BT.RightDistance(idxTrans).*lvconvt);
        data(1:aa,28) = num2cell(MMT_Summary_BT.LeftEdgeCoeff(idxTrans).*lvconvt);
        data(1:aa,29) = num2cell(MMT_Summary_BT.RightEdgeCoeff(idxTrans));
        data(1:aa,30) = num2cell(MMT_Summary_BT.TotalNmbEnsembles(idxTrans));
        data(1:aa,31) = num2cell(MMT_Summary_BT.TotalBadEnsembles(idxTrans));
        data(1:aa,32) = num2cell(MMT_Summary_BT.StartEnsemble(idxTrans));
        data(1:aa,33) = num2cell(MMT_Summary_BT.EndEnsemble(idxTrans));
        data(1:aa,34) = num2cell(MMT_Summary_BT.PercentGoodBins(idxTrans));
        data(1:aa,35) = num2cell(MMT_Summary_BT.PowerCurveCoeff(idxTrans));
        if get(handles.unitsmetric,'Value')
            data(1:aa,36) = num2cell(MMT_Summary_BT.ADCPTemperature(idxTrans));
        else
            data(1:aa,36) = num2cell(((9./5).*MMT_Summary_BT.ADCPTemperature(idxTrans))+32);  
        end
        data(1:aa,37) = num2cell(MMT_Summary_BT.BlankingDistance(idxTrans).*lvconvt);
        data(1:aa,38) = num2cell(MMT_Summary_BT.BinSize(idxTrans).*lvconvt);
        data(1:aa,39) = num2cell(MMT_Summary_BT.BTMode(idxTrans));
        data(1:aa,40) = num2cell(MMT_Summary_BT.WTMode(idxTrans));    
        data(1:aa,41) = num2cell(MMT_Summary_BT.BTPings(idxTrans));
        data(1:aa,42) = num2cell(MMT_Summary_BT.WTPings(idxTrans));       
        data(1:aa,43) = num2cell(MMT_Summary_BT.Begin_Left(idxTrans));
        data(1:aa,44) = num2cell(MMT_Summary_BT.IsSubSectioned(idxTrans));
        data(1:aa,45) = num2cell(MMT_Summary_BT.Use(idxTrans));
        data(1:aa,46) = num2cell(MMT_Active_Config.Offsets_Transducer_Depth(idxTrans).*lvconvt);
        data(1:aa,47) = num2cell(MMT_Active_Config.Proc_BT_Error_Vel_Threshold(idxTrans).*lvconvt);
        data(1:aa,48) = num2cell(MMT_Active_Config.Proc_WT_Error_Vel_Threshold(idxTrans).*lvconvt);	
        data(1:aa,49) = num2cell(MMT_Active_Config.Proc_BT_Up_Vel_Threshold(idxTrans).*lvconvt);
        data(1:aa,50) = num2cell(MMT_Active_Config.Proc_WT_Up_Vel_Threshold(idxTrans).*lvconvt);   
        if isfield(MMT_Active_Config,'Wizard_Commands')
            temp=MMT_Active_Config.Wizard_Commands(1,:);
            ncommands=length(temp);
            wvidx=0;
            woidx=0;
            for j=1:ncommands
                 if cell2mat(strfind(temp(j),'WO'))==1;
                     woidx=j;
                 end
                 if cell2mat(strfind(temp(j),'WV'))==1;
                     wvidx=j;
                 end
            end
            if wvidx>0
                data(1:aa,51) = num2cell(repmat(str2double(strtrim(temp{wvidx}(3:end))),aa,1));
            else
                data(1:aa,51) = num2cell(zeros(aa,1));
            end
            if woidx>0
                commaidx=strfind(temp{woidx},',');
                data(1:aa,52) = num2cell(repmat(str2double(temp{woidx}(3:commaidx-1)),aa,1));
                data(1:aa,53) = num2cell(repmat(str2double(strtrim(temp{woidx}(commaidx+1:end))),aa,1));
            else
                data(1:aa,52) = num2cell(zeros(aa,1));
                data(1:aa,53) = num2cell(zeros(aa,1));
            end
        else
            data(1:aa,51)={''};
            data(1:aa,52)={''};
            data(1:aa,53)={''};
        end
        savefilexls_bt=[prefix '_BT.xlsx'];
        xlswrite (savefilexls_bt,head, 'BT', 'A1');
        xlswrite (savefilexls_bt,data, 'BT', 'A2');  
        
        % Summary for database - dsm
        % --------------------------
        data2{1}=sum(MMT_Summary_BT.Use); % Number of Transects
        data2{2}=data(1,7); % Start Time
        data2{3}=data(data2{1},8); % End Time
        data2{4}=sum([data{:,9}]); % Duration
        data2{5}=mean([data{:,10}]); % Top Q
        data2{6}=100.*std([data{:,10}])./mean([data{:,10}]); % Top Q COV
        data2{7}=mean([data{:,11}]); % Measured Q
        data2{8}=100.*std([data{:,11}])./mean([data{:,11}]); % Measured Q COV
        data2{9}=mean([data{:,12}]); % Bottom Q
        data2{10}=100.*std([data{:,12}])./mean([data{:,12}]); % Bottom Q COV   
        data2{11}=mean([data{:,13}]); % Left Q
        data2{12}=100.*std([data{:,13}])./mean([data{:,13}]); % Left Q COV   
        data2{13}=mean([data{:,14}]); % Right Q
        data2{14}=100.*std([data{:,14}])./mean([data{:,14}]); % Right Q COV 
        data2{15}=mean([data{:,15}]); % Total Q
        data2{16}=100.*std([data{:,15}])./mean([data{:,15}]); % Total Q COV        
        data2{17}=mean([data{:,16}]); % Width
        data2{18}=100.*std([data{:,16}])./mean([data{:,16}]); % Width COV  
        data2{19}=mean([data{:,17}]); % Area
        data2{20}=data2{19}./data2{17}; % Mean Depth
        data2{21}=mean([data{:,36}]); % Temperature
        
    end
    %
    % GGA Reference
    % -------------
    if get(handles.GGAradio,'Value')        
        data(1:aa,1)  = MMT_Summary_GGA.FileName(idxTrans);
        if isnan(MMT_Site_Info.Name)
            data(1:aa,2)={''};
        else
            data(1:aa,2) = cellstr(repmat(MMT_Site_Info.Name,aa,1));
        end
        
        if isnan(MMT_Site_Info.Number)
            data(1:aa,3)={''};
        else
            data(1:aa,3) = cellstr(repmat(MMT_Site_Info.Number,aa,1));
        end
        
        if isnan(MMT_Site_Info.ADCPSerialNmb)
            data(1:aa,4)={''};
        else
            data(1:aa,4) = cellstr(repmat(MMT_Site_Info.ADCPSerialNmb,aa,1));
        end
        data(1:aa,5)  = num2cell(idxTrans);
        data(1:aa,6)  = cellstr(repmat('GGA',aa,1));
        stimeconv=MMT_Summary_GGA.StartTime(idxTrans)./(60*60*24);
        stime=datestr(719529+stimeconv,14);
        data(1:aa,7)  = cellstr(stime);
        etimeconv=MMT_Summary_GGA.EndTime(idxTrans)./(60*60*24);
        etime=datestr(719529+etimeconv,14);
        data(1:aa,8)  = cellstr(etime);
        dursec=MMT_Summary_GGA.EndTime(idxTrans)-MMT_Summary_GGA.StartTime(idxTrans);
        data(1:aa,9)  = num2cell(dursec);
        data(1:aa,10) = num2cell(MMT_Summary_GGA.TopQ(idxTrans).*qconvt);
        data(1:aa,11) = num2cell(MMT_Summary_GGA.MeasuredQ(idxTrans).*qconvt);
        data(1:aa,12) = num2cell(MMT_Summary_GGA.BottomQ(idxTrans).*qconvt);
        data(1:aa,13) = num2cell(MMT_Summary_GGA.LeftQ(idxTrans).*qconvt);
        data(1:aa,14) = num2cell(MMT_Summary_GGA.RightQ(idxTrans).*qconvt);
        data(1:aa,15) = num2cell(MMT_Summary_GGA.TotalQ(idxTrans).*qconvt);
        data(1:aa,16) = num2cell(MMT_Summary_GGA.Width(idxTrans).*lvconvt);
        data(1:aa,17) = num2cell(MMT_Summary_GGA.TotalArea(idxTrans).*aconvt);
        if MMT_Summary_GGA.TotalArea(idxTrans)>0
            data(1:aa,18) =num2cell((MMT_Summary_GGA.TotalQ(idxTrans)./MMT_Summary_GGA.TotalArea(idxTrans)).*lvconvt);
        end
        data(1:aa,19) = num2cell(MMT_Summary_GGA.FlowDirection(idxTrans));
        data(1:aa,20) = num2cell(MMT_Summary_GGA.MaxWaterSpeed(idxTrans).*lvconvt);
        data(1:aa,21) = num2cell(MMT_Summary_GGA.MeanRiverVel(idxTrans).*lvconvt);
        data(1:aa,22) = num2cell(MMT_Summary_GGA.MaxWaterDepth(idxTrans).*lvconvt);
        data(1:aa,23) = num2cell(MMT_Summary_GGA.MeanWaterDepth(idxTrans).*lvconvt);
        data(1:aa,24) = num2cell(MMT_Summary_GGA.MeanBoatSpeed(idxTrans).*lvconvt);
        data(1:aa,25) = num2cell(MMT_Summary_GGA.MeanBoatCourse(idxTrans));
        data(1:aa,26) = num2cell(MMT_Summary_GGA.LeftDistance(idxTrans).*lvconvt);
        data(1:aa,27) = num2cell(MMT_Summary_GGA.RightDistance(idxTrans).*lvconvt);
        data(1:aa,28) = num2cell(MMT_Summary_GGA.LeftEdgeCoeff(idxTrans));
        data(1:aa,29) = num2cell(MMT_Summary_GGA.RightEdgeCoeff(idxTrans));
        data(1:aa,30) = num2cell(MMT_Summary_GGA.TotalNmbEnsembles(idxTrans));
        data(1:aa,31) = num2cell(MMT_Summary_GGA.TotalBadEnsembles(idxTrans));
        data(1:aa,32) = num2cell(MMT_Summary_GGA.StartEnsemble(idxTrans));
        data(1:aa,33) = num2cell(MMT_Summary_GGA.EndEnsemble(idxTrans));
        data(1:aa,34) = num2cell(MMT_Summary_GGA.PercentGoodBins(idxTrans));
        data(1:aa,35) = num2cell(MMT_Summary_GGA.PowerCurveCoeff(idxTrans));
        if get(handles.unitsmetric,'Value')
            data(1:aa,36) = num2cell(MMT_Summary_GGA.ADCPTemperature(idxTrans));
        else
            data(1:aa,36) = num2cell(((9./5).*MMT_Summary_GGA.ADCPTemperature(idxTrans))+32);  
        end
        data(1:aa,37) = num2cell(MMT_Summary_GGA.BlankingDistance(idxTrans).*lvconvt);
        data(1:aa,38) = num2cell(MMT_Summary_GGA.BinSize(idxTrans).*lvconvt);
        data(1:aa,39) = num2cell(MMT_Summary_GGA.BTMode(idxTrans));
        data(1:aa,40) = num2cell(MMT_Summary_GGA.WTMode(idxTrans));    
        data(1:aa,41) = num2cell(MMT_Summary_GGA.BTPings(idxTrans));
        data(1:aa,42) = num2cell(MMT_Summary_GGA.WTPings(idxTrans));       
        data(1:aa,43) = num2cell(MMT_Summary_GGA.Begin_Left(idxTrans));
        data(1:aa,44) = num2cell(MMT_Summary_GGA.IsSubSectioned(idxTrans));
        data(1:aa,45) = num2cell(MMT_Summary_GGA.Use(idxTrans));
        data(1:aa,46) = num2cell(MMT_Active_Config.Offsets_Transducer_Depth(idxTrans).*lvconvt);
        data(1:aa,47) = num2cell(MMT_Active_Config.Proc_BT_Error_Vel_Threshold(idxTrans).*lvconvt);
        data(1:aa,48) = num2cell(MMT_Active_Config.Proc_WT_Error_Vel_Threshold(idxTrans).*lvconvt);	
        data(1:aa,49) = num2cell(MMT_Active_Config.Proc_BT_Up_Vel_Threshold(idxTrans).*lvconvt);
        data(1:aa,50) = num2cell(MMT_Active_Config.Proc_WT_Up_Vel_Threshold(idxTrans).*lvconvt);   
        if isfield(MMT_Active_Config,'Wizard_Commands')
            temp=MMT_Active_Config.Wizard_Commands(1,:);
            ncommands=length(temp);
            wvidx=0;
            woidx=0;
            for j=1:ncommands
                 if cell2mat(strfind(temp(j),'WO'))==1;
                     woidx=j;
                 end
                 if cell2mat(strfind(temp(j),'WV'))==1;
                     wvidx=j;
                 end
            end
            if wvidx>0
                data(1:aa,51) = num2cell(repmat(str2double(strtrim(temp{wvidx}(3:end))),aa,1));
            else
                data(1:aa,51) = num2cell(zeros(aa,1));
            end
            if woidx>0
                commaidx=strfind(temp{woidx},',');
                data(1:aa,52) = num2cell(repmat(str2double(temp{woidx}(3:commaidx-1)),aa,1));
                data(1:aa,53) = num2cell(repmat(str2double(strtrim(temp{woidx}(commaidx+1:end))),aa,1));
            else
                data(1:aa,52) = num2cell(zeros(aa,1));
                data(1:aa,53) = num2cell(zeros(aa,1));
            end
        else
            data(1:aa,51)={''};
            data(1:aa,52)={''};
            data(1:aa,53)={''};
        end     
        savefilexls_gga=[prefix '_GGA.xlsx'];
        xlswrite (savefilexls_gga,head, 'GGA', 'A1');
        xlswrite (savefilexls_gga,data, 'GGA', 'A2');
        
        % Summary for database - dsm
        % --------------------------
        data2{1}=sum(MMT_Summary_GGA.Use); % Number of Transects
        data2{2}=data(1,7); % Start Time
        data2{3}=data(data2{1},8); % End Time
        data2{4}=sum([data{:,9}]); % Duration
        data2{5}=mean([data{:,10}]); % Top Q
        data2{6}=100.*std([data{:,10}])./mean([data{:,10}]); % Top Q COV
        data2{7}=mean([data{:,11}]); % Measured Q
        data2{8}=100.*std([data{:,11}])./mean([data{:,11}]); % Measured Q COV
        data2{9}=mean([data{:,12}]); % Bottom Q
        data2{10}=100.*std([data{:,12}])./mean([data{:,12}]); % Bottom Q COV   
        data2{11}=mean([data{:,13}]); % Left Q
        data2{12}=100.*std([data{:,13}])./mean([data{:,13}]); % Left Q COV   
        data2{13}=mean([data{:,14}]); % Right Q
        data2{14}=100.*std([data{:,14}])./mean([data{:,14}]); % Right Q COV 
        data2{15}=mean([data{:,15}]); % Total Q
        data2{16}=100.*std([data{:,15}])./mean([data{:,15}]); % Total Q COV        
        data2{17}=mean([data{:,16}]); % Width
        data2{18}=100.*std([data{:,16}])./mean([data{:,16}]); % Width COV  
        data2{19}=mean([data{:,17}]); % Area
        data2{20}=data2{19}./data2{17}; % Mean Depth
        data2{21}=mean([data{:,36}]); % Temperature        
    end
    %
    % VTG Reference
    % -------------
    if get(handles.VTGradio,'Value')      
        data(1:aa,1)  = MMT_Summary_VTG.FileName(idxTrans);
        if isnan(MMT_Site_Info.Name)
            data(1:aa,2)={''};
        else
            data(1:aa,2) = cellstr(repmat(MMT_Site_Info.Name,aa,1));
        end
        
        if isnan(MMT_Site_Info.Number)
            data(1:aa,3)={''};
        else
            data(1:aa,3) = cellstr(repmat(MMT_Site_Info.Number,aa,1));
        end
        
        if isnan(MMT_Site_Info.ADCPSerialNmb)
            data(1:aa,4)={''};
        else
            data(1:aa,4) = cellstr(repmat(MMT_Site_Info.ADCPSerialNmb,aa,1));
        end
        data(1:aa,5)  = num2cell(idxTrans);
        data(1:aa,6)  = cellstr(repmat('VTG',aa,1));
        stimeconv=MMT_Summary_VTG.StartTime(idxTrans)./(60*60*24);
        stime=datestr(719529+stimeconv,14);
        data(1:aa,7)  = cellstr(stime);
        etimeconv=MMT_Summary_VTG.EndTime(idxTrans)./(60*60*24);
        etime=datestr(719529+etimeconv,14);
        data(1:aa,8)  = cellstr(etime);
        dursec=MMT_Summary_VTG.EndTime(idxTrans)-MMT_Summary_VTG.StartTime(idxTrans);
        data(1:aa,9)  = num2cell(dursec);
        data(1:aa,10) = num2cell(MMT_Summary_VTG.TopQ(idxTrans).*qconvt);
        data(1:aa,11) = num2cell(MMT_Summary_VTG.MeasuredQ(idxTrans).*qconvt);
        data(1:aa,12) = num2cell(MMT_Summary_VTG.BottomQ(idxTrans).*qconvt);
        data(1:aa,13) = num2cell(MMT_Summary_VTG.LeftQ(idxTrans).*qconvt);
        data(1:aa,14) = num2cell(MMT_Summary_VTG.RightQ(idxTrans).*qconvt);
        data(1:aa,15) = num2cell(MMT_Summary_VTG.TotalQ(idxTrans).*qconvt);
        data(1:aa,16) = num2cell(MMT_Summary_VTG.Width(idxTrans).*lvconvt);
        data(1:aa,17) = num2cell(MMT_Summary_VTG.TotalArea(idxTrans).*aconvt);
        if MMT_Summary_VTG.TotalArea(idxTrans)>0
            data(1:aa,18) =num2cell((MMT_Summary_VTG.TotalQ(idxTrans)./MMT_Summary_VTG.TotalArea(idxTrans)).*lvconvt);
        end
        data(1:aa,19) = num2cell(MMT_Summary_VTG.FlowDirection(idxTrans));
        data(1:aa,20) = num2cell(MMT_Summary_VTG.MaxWaterSpeed(idxTrans).*lvconvt);
        data(1:aa,21) = num2cell(MMT_Summary_VTG.MeanRiverVel(idxTrans).*lvconvt);
        data(1:aa,22) = num2cell(MMT_Summary_VTG.MaxWaterDepth(idxTrans).*lvconvt);
        data(1:aa,23) = num2cell(MMT_Summary_VTG.MeanWaterDepth(idxTrans).*lvconvt);
        data(1:aa,24) = num2cell(MMT_Summary_VTG.MeanBoatSpeed(idxTrans).*lvconvt);
        data(1:aa,25) = num2cell(MMT_Summary_VTG.MeanBoatCourse(idxTrans));
        data(1:aa,26) = num2cell(MMT_Summary_VTG.LeftDistance(idxTrans).*lvconvt);
        data(1:aa,27) = num2cell(MMT_Summary_VTG.RightDistance(idxTrans).*lvconvt);
        data(1:aa,28) = num2cell(MMT_Summary_VTG.LeftEdgeCoeff(idxTrans));
        data(1:aa,29) = num2cell(MMT_Summary_VTG.RightEdgeCoeff(idxTrans));
        data(1:aa,30) = num2cell(MMT_Summary_VTG.TotalNmbEnsembles(idxTrans));
        data(1:aa,31) = num2cell(MMT_Summary_VTG.TotalBadEnsembles(idxTrans));
        data(1:aa,32) = num2cell(MMT_Summary_VTG.StartEnsemble(idxTrans));
        data(1:aa,33) = num2cell(MMT_Summary_VTG.EndEnsemble(idxTrans));
        data(1:aa,34) = num2cell(MMT_Summary_VTG.PercentGoodBins(idxTrans));
        data(1:aa,35) = num2cell(MMT_Summary_VTG.PowerCurveCoeff(idxTrans));
        if get(handles.unitsmetric,'Value')
            data(1:aa,36) = num2cell(MMT_Summary_VTG.ADCPTemperature(idxTrans));
        else
            data(1:aa,36) = num2cell(((9./5).*MMT_Summary_VTG.ADCPTemperature(idxTrans))+32);  
        end
        data(1:aa,37) = num2cell(MMT_Summary_VTG.BlankingDistance(idxTrans).*lvconvt);
        data(1:aa,38) = num2cell(MMT_Summary_VTG.BinSize(idxTrans).*lvconvt);
        data(1:aa,39) = num2cell(MMT_Summary_VTG.BTMode(idxTrans));
        data(1:aa,40) = num2cell(MMT_Summary_VTG.WTMode(idxTrans));    
        data(1:aa,41) = num2cell(MMT_Summary_VTG.BTPings(idxTrans));
        data(1:aa,42) = num2cell(MMT_Summary_VTG.WTPings(idxTrans));       
        data(1:aa,43) = num2cell(MMT_Summary_VTG.Begin_Left(idxTrans));
        data(1:aa,44) = num2cell(MMT_Summary_VTG.IsSubSectioned(idxTrans));
        data(1:aa,45) = num2cell(MMT_Summary_VTG.Use(idxTrans));
        data(1:aa,46) = num2cell(MMT_Active_Config.Offsets_Transducer_Depth(idxTrans).*lvconvt);
        data(1:aa,47) = num2cell(MMT_Active_Config.Proc_BT_Error_Vel_Threshold(idxTrans).*lvconvt);
        data(1:aa,48) = num2cell(MMT_Active_Config.Proc_WT_Error_Vel_Threshold(idxTrans).*lvconvt);	
        data(1:aa,49) = num2cell(MMT_Active_Config.Proc_BT_Up_Vel_Threshold(idxTrans).*lvconvt);
        data(1:aa,50) = num2cell(MMT_Active_Config.Proc_WT_Up_Vel_Threshold(idxTrans).*lvconvt);   
        if isfield(MMT_Active_Config,'Wizard_Commands')
            temp=MMT_Active_Config.Wizard_Commands(1,:);
            ncommands=length(temp);
            wvidx=0;
            woidx=0;
            for j=1:ncommands
                 if cell2mat(strfind(temp(j),'WO'))==1;
                     woidx=j;
                 end
                 if cell2mat(strfind(temp(j),'WV'))==1;
                     wvidx=j;
                 end
            end
            if wvidx>0
                data(1:aa,51) = num2cell(repmat(str2double(strtrim(temp{wvidx}(3:end))),aa,1));
            else
                data(1:aa,51) = num2cell(zeros(aa,1));
            end
            if woidx>0
                commaidx=strfind(temp{woidx},',');
                data(1:aa,52) = num2cell(repmat(str2double(temp{woidx}(3:commaidx-1)),aa,1));
                data(1:aa,53) = num2cell(repmat(str2double(strtrim(temp{woidx}(commaidx+1:end))),aa,1));
            else
                data(1:aa,52) = num2cell(zeros(aa,1));
                data(1:aa,53) = num2cell(zeros(aa,1));
            end
        else
            data(1:aa,51)={''};
            data(1:aa,52)={''};
            data(1:aa,53)={''};
        end         
        savefilexls_vtg=[prefix '_VTG.xlsx'];
        xlswrite (savefilexls_vtg,head, 'VTG', 'A1');
        xlswrite (savefilexls_vtg,data, 'VTG', 'A2');
        
        % Summary for database - dsm
        % --------------------------
        data2{1}=sum(MMT_Summary_VTG.Use); % Number of Transects
        data2{2}=data(1,7); % Start Time
        data2{3}=data(data2{1},8); % End Time
        data2{4}=sum([data{:,9}]); % Duration
        data2{5}=mean([data{:,10}]); % Top Q
        data2{6}=100.*std([data{:,10}])./mean([data{:,10}]); % Top Q COV
        data2{7}=mean([data{:,11}]); % Measured Q
        data2{8}=100.*std([data{:,11}])./mean([data{:,11}]); % Measured Q COV
        data2{9}=mean([data{:,12}]); % Bottom Q
        data2{10}=100.*std([data{:,12}])./mean([data{:,12}]); % Bottom Q COV   
        data2{11}=mean([data{:,13}]); % Left Q
        data2{12}=100.*std([data{:,13}])./mean([data{:,13}]); % Left Q COV   
        data2{13}=mean([data{:,14}]); % Right Q
        data2{14}=100.*std([data{:,14}])./mean([data{:,14}]); % Right Q COV 
        data2{15}=mean([data{:,15}]); % Total Q
        data2{16}=100.*std([data{:,15}])./mean([data{:,15}]); % Total Q COV        
        data2{17}=mean([data{:,16}]); % Width
        data2{18}=100.*std([data{:,16}])./mean([data{:,16}]); % Width COV  
        data2{19}=mean([data{:,17}]); % Area
        data2{20}=data2{19}./data2{17}; % Mean Depth
        data2{21}=mean([data{:,36}]); % Temperature        
    end     
    %
    % Processing complete reenable pushbutton
    % ---------------------------------------
    set(handles.pushbutton1,'Enable','on')
    set(handles.status,'String','Excel File Created')
    drawnow;
    %
    % Update handles structure
    % ------------------------
    guidata(hObject, handles);
end
%
% If an error occurs report to user
% ----------------------------------
catch ME
	errordlg(ME.message)
    set(handles.pushbutton1,'Enable','on')
    drawnow;
end


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear
close all


% --------------------------------------------------------------------
function Mode_pannel_ButtonDownFcn(hObject, eventdata, handles)
% hObject    handle to Mode_pannel (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on key press with focus on GGAradio and none of its controls.
function GGAradio_KeyPressFcn(hObject, eventdata, handles)
% hObject    handle to GGAradio (see GCBO)
% eventdata  structure with the following fields (see UICONTROL)
%	Key: name of the key that was pressed, in lower case
%	Character: character interpretation of the key(s) that was pressed
%	Modifier: name(s) of the modifier key(s) (i.e., control, shift) pressed
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on mouse press over figure background, over a disabled or
% --- inactive control, or over an axes background.
function figure1_WindowButtonUpFcn(hObject, eventdata, handles)
% hObject    handle to figure1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pbReprocess.
function pbReprocess_Callback(hObject, eventdata, handles)
% hObject    handle to pbReprocess (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
    handles.reprocess=1;
    %
    % Update handles structure
    % ------------------------
    guidata(hObject, handles);
    pushbutton1_Callback(hObject, eventdata, handles)


% --- Executes on button press in BTradio.
function BTradio_Callback(hObject, eventdata, handles)
% hObject    handle to BTradio (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of BTradio


% --- Executes on button press in GGAradio.
function GGAradio_Callback(hObject, eventdata, handles)
% hObject    handle to GGAradio (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of GGAradio


% --- Executes on button press in VTGradio.
function VTGradio_Callback(hObject, eventdata, handles)
% hObject    handle to VTGradio (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of VTGradio
