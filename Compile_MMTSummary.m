% Compile_MMTSummary.m
% Run this script to compile MMTSummary on a local machine running Matlab.
% Assumes all components are in the present working directory.
% 
% Frank L. Engel, USGS, IL WSC


% Destination of EXE
savefile = pwd;

% Command string
com_str = ['-o MMTSummary -W WinMain -T link:exe -d '...
    savefile ...
    ' -N -v '...
    'MMTSummary.m -a MMTSummary.fig'];

% Compile
eval(['mcc ' com_str])


% mcc -o VMT -W WinMain -T link:exe -d C:\Users\fengel\Downloads\VMT\src -N
% -p map -p stats -p images -p utils -p docs -p tools -v
% D:\Scratch\VMT_fle_merge\VMT.m -a D:\Scratch\VMT_fle_merge\VMT.fig