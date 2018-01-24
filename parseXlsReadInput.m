function [sheet,x1Range,processFcn,basic] = parseXlsReadInput(varargin)
% Parse the input of the xlread function to obtain all relevant outputs for
% the input of an xls/xlsx file (independent on matlab implementations).

sheet = 1;
x1Range = '';
processFcn = [];
basic = 0;
if numel(varargin) == 0
    return;
end

sheetDone = false;
rangeDone = false;
basicDone = false;

varargpos = 1;

while varargpos <= numel(varargin)
    carg = varargin{varargpos};
    if basicDone %Now, we have to be in the last element, which is the process function
        processFcn = carg;
        return;
    end
    if ~sheetDone
        if isnumeric(carg) || ischar(carg) && numel(strsplit(carg,':')) < 2 
            sheet = carg;
            sheetDone = true;
        else %its not numeric, and looks like a x1Range
            x1Range = carg;
            %We did not yet parse the sheet, but got a range -> only range
            %provided.
            break;
        end
    else %if sheet is done, we got a second argument, which has to be range.
        if ~rangeDone
            x1Range = carg;
            rangeDone = true;
        else
            basic = strcmp('basic',carg);
            basicDone = true;
        end
    end

    varargpos = varargpos + 1;    
end