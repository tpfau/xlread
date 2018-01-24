function setupxlread()
%setupxlread sets up all folders and pathes necessary to use xlread.
% 
% USAGE:
%
%    setupxlwrite()
% .. Author: - Thomas Pfau Jan 2018

if exist('org.apache.poi.ss.usermodel.WorkbookFactory', 'class') ~= 8     
    oldfolder = pwd;
    folder = fileparts(which('xlread'));
    cd(folder);
    addpath('poi_library');
    javaaddpath(['poi_library' filesep 'poi-3.8-20120326.jar']);
    javaaddpath(['poi_library' filesep 'poi-ooxml-3.8-20120326.jar']);
    javaaddpath(['poi_library' filesep 'poi-ooxml-schemas-3.8-20120326.jar']);
    javaaddpath(['poi_library' filesep 'xmlbeans-2.3.0.jar']);
    javaaddpath(['poi_library' filesep 'dom4j-1.6.1.jar']);
    cd(oldfolder);
end