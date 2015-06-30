% Run MMTSummary Recurrsively


% Set Options
guiparams.metric_units              = false;
guiparams.bottom_track_reference    = true;
guiparams.gga_reference             = false;
guiparams.vtg_reference             = false;

% Get a list of files to work with
dname = uigetdir(pwd,'Select directory containing mmt files (will search recursively):');
[~,~,files] = dirr([dname filesep '*.mmt'],'name');

% Loop through each file and process
for i = 1:length(files)
    [inpath, infile, ext] = fileparts(files{i});
    MMTProcessingEngine([inpath filesep],[infile ext],guiparams)
    movefile(fullfile(inpath,[infile '.xlsx']),[dname filesep])
end