function [num,txt,raw,costum] = xlread(filename,varargin)
% XLREAD reads a microsoft xls or xlsx file using the POI library.
%   The syntax is the same as xlsread from Matlab 2017b
% processFcn only supports the Value and Count fields of the Data object
%            otherwise passed in to the function. 
%==============================================================================
% Author: Thomas Pfau Jan 2018

% Check if POI lib is loaded and if not, load it.

while exist('org.apache.poi.ss.usermodel.WorkbookFactory', 'class') ~= 8                
    setupxlread();
end

% Import required POI Java Classes
import org.apache.poi.ss.usermodel.*;

[sheet,range,processFcn,basic] = parseXlsReadInput(varargin{:});

% Open a file
xlsFile = java.io.File(filename);
num = [];
txt = {};
raw = {};

% If file does not exist create a new workbook
if xlsFile.isFile()
    % create XSSF or HSSF workbook from existing workbook
    fileIn = java.io.FileInputStream(xlsFile);
    xlsWorkbook = WorkbookFactory.create(fileIn);
else
    error('File %s not found',filename);
end

%Read from the given sheet.
if ~isempty(sheet)
    if isnumeric(sheet)        
        % Use Sheet -1 as POI is 0 based, and matlab is 1-based.
        if xlsWorkbook.getNumberOfSheets() >= sheet && sheet >= 1
            xlsSheet = xlsWorkbook.getSheetAt(sheet-1);
        else
            error('The Excel file only has %i sheets while sheet %i as requested.',xlsWorkbook.getNumberOfSheets(), sheet);
        end
    else
        %If its a name, we will first have to collect the sheet names:
        sheetNames = cell(1,xlsWorkbook.getNumberOfSheets());
        for i = 1:size(sheetNames,2)
            sheetNames{i} = xlsWorkbook.getSheetAt(i-1).getSheetName();
        end
        sheetIndex = find(cellfun(@(x) strcmpi(x,sheet),sheetNames));
        xlsSheet = xlsWorkbook.getSheetAt(sheetIndex-1);
    end
else
    % check number of sheets
    nSheets = xlsWorkbook.getNumberOfSheets();
    
    % If no sheets, return empty data
    if nSheets < 1
        return
    else
        xlsSheet = xlsWorkbook.getSheetAt(0);
    end
end

%Now, we got the requested XLS sheet.

if isempty(range)
    iRowStart = 0;
    iColStart = 0;
    iRowEnd = xlsSheet.getLastRowNum();
    iColEnd = inf;
    %We will read everything.
else
    % Define start & end cell
    cellStart = range(1:iSeperator-1);
    cellEnd = range(iSeperator+1:end);
    
    % Create a helper to get the row and column
    cellStart = CellReference(cellStart);
    cellEnd = CellReference(cellEnd);
    
    % Get start & end locations
    iRowStart = cellStart.getRow();
    iColStart = cellStart.getCol();
    iRowEnd = cellEnd.getRow();
    iColEnd = cellEnd.getCol();
end
selCols = (iColEnd - iColStart) + 1;
selRows = (iRowEnd - iRowStart) + 1;

numCols = 0;
%get the maximal number of cells in a row
for i = 0:xlsSheet.getLastRowNum()
    numCols = max(numCols,xlsSheet.getRow(i).getLastCellNum());
end
%Lets get a rough estimation on the size of the sheet in order to
%initialize our outputs.

numCols = min(selCols,numCols);
numRows = min(selRows,xlsSheet.getLastRowNum()+1);

raw = cell(numRows,numCols);
%raw(:) = {NaN}; %Default is NaN.

% Iterate over all data
for iRow = iRowStart:min(iRowEnd)
    % Fetch the row (if it exists)
    currentRow = xlsSheet.getRow(iRow);
    % enter data for all cols
    for iCol = iColStart:min(iColEnd,currentRow.getLastCellNum())
        % Check if cell exists
        currentCell = currentRow.getCell(iCol);
        if ~isempty(currentCell) %No information, pass
            switch currentCell.getCellType()
                case {currentCell.CELL_TYPE_NUMERIC,currentCell.CELL_TYPE_BOOLEAN}
                    raw{iRow+1,iCol+1} = currentCell.getNumericCellValue();
                case currentCell.CELL_TYPE_STRING
                    raw{iRow+1,iCol+1} = char(currentCell.getStringCellValue());
                case currentCell.CELL_TYPE_FORMULA
                    %This is a bit more interesting.
                    switch currentCell.getCachedFormulaResultType
                        case currentCell.CELL_TYPE_STRING
                            raw{iRow+1,iCol+1} = char(currentCell.getStringCellValue());
                        case {currentCell.CELL_TYPE_NUMERIC,currentCell.CELL_TYPE_BOOLEAN}
                            raw{iRow+1,iCol+1} = double(currentCell.getNumericCellValue());
                    end
            end
        end
    end
end

%Anything that is empty, will become NaN.
raw(cellfun(@isempty, raw)) = {NaN};

if ~isempty(processFcn)
    Data.Value = raw;
    Data.Count = numel(raw);    
    try
        [raw,costum] = processFcn(Data);
    catch %probably not two outputs...
        [raw] = processFcn(Data);
    end
end

[x,y] = find(cellfun(@(x) ischar(x) && ~isempty(x),raw));
xmin = min(x);
xmax = max(x);
ymin = min(y);
ymax = max(y);
txt = raw(xmin:xmax,ymin:ymax);
txt(cellfun(@(x) ~ischar(x) | isempty(x),txt)) = {''};

[x,y] = find(cellfun(@(x) isnumeric(x) && ~isempty(x) && ~isnan(x),raw));
xmin = min(x);
xmax = max(x);
ymin = min(y);
ymax = max(y);
num = raw(xmin:xmax,ymin:ymax);
num(cellfun(@(x) ~isnumeric(x) | isempty(x) ,num)) = {NaN};
num = cell2mat(num);

fileIn.close();

end
