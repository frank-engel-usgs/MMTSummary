function varargout = MMTSummary(varargin)
% --- MMTSummary ---
% This is a GUI which will create or overwrite an Excel Spreadsheet
% containing the summary discharge information from a Win River II
% Measurement file (*.MMT). The Excel spreadsheet result is similar to the
% discharge tablature (F12) in WRII. 
% 
%
% See also: mmt2mat, parseXML
% 
% Code originally written by:
%    Dave Mueller
% with contributions from:
%    Jeremiah ???
%    Justin Boldt
%    Frank Engel
% 
% GUI design and coding layout modified by
%    Frank L. Engel, USGS, IL WSC, Mar 3, 2013

% Last Modified by GUIDE v2.5 07-Mar-2013 16:21:22



% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @MMTSummary_OpeningFcn, ...
                   'gui_OutputFcn',  @MMTSummary_OutputFcn, ...
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


% --------------------------------------------------------------------
function MMTSummary_OpeningFcn(hObject, eventdata, handles, varargin)

% Load the GUI preferences:
% -------------------------
load_prefs(handles.figure1)

% Initialize the GUI parameters:
% ------------------------------
guiparams.bottom_track_reference    = 1;
guiparams.gga_reference             = 0;
guiparams.vtg_reference             = 0;
guiparams.metric_units              = 1;
guiparams.english_units             = 0;
guiparams.status_message_box        = 'No MMT file loaded';
guiparams.file                      = '';
guiparams.mmtpath                   = '';
guiparams.version                   = 'v2.0';

% Store the application data:
% ---------------------------
setappdata(handles.figure1,'guiparams',guiparams)

% Initialize the GUI:
% -------------------
initGUI(handles)

% UIWAIT makes MMTSummary wait for user response (see UIRESUME)
% uiwait(handles.figure1);

% {EOF] MMTSummary_OpeningFcn


% --------------------------------------------------------------------
function varargout = MMTSummary_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles;

% [EOF] MMTSummary_OutputFcn


% --------------------------------------------------------------------
function LoadMMT_Callback(hObject, eventdata, handles)
 
% Get the application data
% ------------------------
guiparams = getappdata(handles.figure1,'guiparams');
guiprefs  = getappdata(handles.figure1,'guiprefs');

% Get file
% --------
[guiparams.file,guiparams.mmtpath] = ...
    uigetfile({'*.mmt','All .mmt Files'; '*.*','All Files'},'Select .mmt File',[guiprefs.mmtpath guiprefs.file]);

% If user cancels uigetfile, set file and path
if any(guiparams.mmtpath==0) || any(guiparams.file==0)
    guiparams.file    = '';
    guiparams.mmtpath = pwd;
    
    % Update status window
    set(handles.StatusMessageBox, 'String', 'No MMT file loaded')
else
    % Update status window
    guiparams.status_message_box = ['Current file: ' guiparams.file];
    set(handles.StatusMessageBox, 'String', guiparams.status_message_box)
end
    
% Store the application data
% --------------------------
setappdata(handles.figure1,'guiparams',guiparams);

% Update the preferences:
% -----------------------
guiprefs = getappdata(handles.figure1,'guiprefs');
guiprefs.mmtpath = guiparams.mmtpath;
guiprefs.file = guiparams.file;
setappdata(handles.figure1,'guiprefs',guiprefs)
store_prefs(handles.figure1,'mmtpath')
store_prefs(handles.figure1,'file')

% [EOF] LoadMMT_Callback





% --------------------------------------------------------------------
function Process_Callback(hObject, eventdata, handles)

%     handles.reprocess=1;
%     %
%     % Update handles structure
%     % ------------------------
%     guidata(hObject, handles);
%     pushbutton1_Callback(hObject, eventdata, handles)
    
% Get the application data
% ------------------------
guiparams = getappdata(handles.figure1,'guiparams');
guiprefs = getappdata(handles.figure1,'guiprefs');

% Process the file
% ----------------
MMTProcessingEngine(guiprefs.mmtpath,guiprefs.file,guiparams)

% Store the application data
% --------------------------
setappdata(handles.figure1,'guiparams',guiparams);
% [EOF] Process_Callback


% --------------------------------------------------------------------
function Close_Callback(hObject, eventdata, handles)
% hObject    handle to Close (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear
close all hidden

% [EOF] Close_Callback


% --------------------------------------------------------------------
function BottomTrackReference_Callback(hObject, eventdata, handles)
% Get the application data
% ------------------------
guiparams = getappdata(handles.figure1,'guiparams');

guiparams.bottom_track_reference = get(handles.BottomTrackReference,'Value');

% Store the application data
% --------------------------
setappdata(handles.figure1,'guiparams',guiparams);


% --------------------------------------------------------------------
function GGAReference_Callback(hObject, eventdata, handles)
% Get the application data
% ------------------------
guiparams = getappdata(handles.figure1,'guiparams');

guiparams.gga_reference = get(handles.GGAReference,'Value');

% Store the application data
% --------------------------
setappdata(handles.figure1,'guiparams',guiparams);

% --------------------------------------------------------------------
function VTGReference_Callback(hObject, eventdata, handles)
% Get the application data
% ------------------------
guiparams = getappdata(handles.figure1,'guiparams');

guiparams.vtg_reference = get(handles.VTGReference,'Value');

% Store the application data
% --------------------------
setappdata(handles.figure1,'guiparams',guiparams);

% --- Executes on button press in MetricUnits.
function MetricUnits_Callback(hObject, eventdata, handles)
% Get the application data
% ------------------------
guiparams = getappdata(handles.figure1,'guiparams');

if get(handles.MetricUnits,'Value')
    set(handles.MetricUnits, 'Value',1)
    set(handles.EnglishUnits,'Value',0)
else
    set(handles.MetricUnits, 'Value',0)
    set(handles.EnglishUnits,'Value',1)
end

guiparams.metric_units  = get(handles.MetricUnits, 'Value');
guiparams.english_units = get(handles.EnglishUnits,'Value');

% Store the application data
% --------------------------
setappdata(handles.figure1,'guiparams',guiparams);


% --- Executes on button press in EnglishUnits.
function EnglishUnits_Callback(hObject, eventdata, handles)
% Get the application data
% ------------------------
guiparams = getappdata(handles.figure1,'guiparams');

if get(handles.EnglishUnits,'Value')
    set(handles.MetricUnits, 'Value',0)
    set(handles.EnglishUnits,'Value',1)
else
    set(handles.MetricUnits, 'Value',1)
    set(handles.EnglishUnits,'Value',0)
end

guiparams.metric_units  = get(handles.MetricUnits, 'Value');
guiparams.english_units = get(handles.EnglishUnits,'Value');

% Store the application data
% --------------------------
setappdata(handles.figure1,'guiparams',guiparams);

%%%%%%%%%%%%%%%%
% SUBFUNCTIONS %
%%%%%%%%%%%%%%%%

% --------------------------------------------------------------------
function initGUI(handles)
% Initialize the UI controls in the GUI.

% Get the application data
% ------------------------
guiparams = getappdata(handles.figure1,'guiparams');

% Set the name and version
% ------------------------
set(handles.figure1,'Name',['MMT to Excel Utility ' guiparams.version], ...
    'DockControls','off')

% Set reference panel up
% ----------------------
set(handles.BottomTrackReference,'Value',guiparams.bottom_track_reference)
set(handles.GGAReference,        'Value',guiparams.gga_reference)
set(handles.VTGReference,        'Value',guiparams.vtg_reference)

% Set units panel up
% ------------------
if guiparams.metric_units
    set(handles.MetricUnits, 'Value', 1)
    set(handles.EnglishUnits,'Value', 0)
else
    set(handles.MetricUnits, 'Value', 0)
    set(handles.EnglishUnits,'Value', 1)
end

% Set status message
% ------------------
set(handles.StatusMessageBox,'String',guiparams.status_message_box)

% [EOF] initGUI

% --------------------------------------------------------------------
function set_enable(handles,enable_state)
% Modify what is enabled in the GUI based on runtime state

% Get the application data
% ------------------------
guiparams = getappdata(handles.figure1,'guiparams');

switch enable_state
    case 'init'
        set([handles.LoadMMT
            handles.Process], 'Enable', 'off')
    case 'fileloaded'
        set([handles.LoadMMT
            handles.Process], 'Enable', 'on')
end

% [EOF] set_enable




% --------------------------------------------------------------------
function load_prefs(hfigure)
% Load the GUI preferences.  Also, initialize preferences if they don't
% already exist.

% Preferences:
% 'mmtpath'             Path of current MMT file
% 'file'                filename of current MMT file
% 
% Groups:
% 'MMT'

% mmtpath
if ispref('MMT','mmtpath')
    mmt = getpref('MMT','mmtpath');
    if exist(mmt.mmtpath,'dir') % does the directory exist?
        guiprefs.mmtpath = mmt.mmtpath;
    else
        guiprefs.mmtpath = pwd;
    end

else % Initialize pref
    guiprefs.mmtpath = pwd;
    mmt.mmtpath = pwd;
    mmt.file = '';
    setpref('MMT','mmtpath',mmt)
end

% file
if ispref('MMT','file')
    file = getpref('MMT','file');
    if exist(fullfile(mmt.mmtpath,file.file),'file')==2
        guiprefs.file = file.file;
    else
        guiprefs.file = '';
    end
else % Initialize pref
    guiprefs.file = '';
    file.file     = '';
    setpref('MMT','file',file)
end

% Store application data
% ----------------------
setappdata(hfigure,'guiprefs',guiprefs)

% [EOF] load_prefs

% --------------------------------------------------------------------
function store_prefs(hfigure,pref)
% Store preferences in the Application data and in the persistent
% preferences data.
% 
% NOTE: This eliminates need for LastDir.mat setup. All prefs are
% persistent as long as the MCR (or Matlab) is running.

% Preferences:
% 'mmtpath'             Path of current MMT file
% 'file'                filename of current MMT file
% 
% Groups:
% 'MMT'

% Get the current preferences
% ---------------------------
guiprefs = getappdata(hfigure,'guiprefs');

% Push prefs to persistent memory
% -------------------------------
switch pref
    case 'mmtpath'
        mmt.mmtpath = guiprefs.mmtpath;
        setpref('MMT','mmtpath',mmt)
    case 'file'
        mmt.file = guiprefs.file;
        setpref('MMT','file',mmt)
    otherwise
end

% [EOF] store_prefs



