# xlread

`xlread` offers `xlsread` functionality for unix machines and Windows machines without excel with the same syntax as xlsread without the need of an excel backend.
`xlread` offers non Excel systems the option to use the processFcn argument of `xlsread`, but is currently restricted to the Data.Value and Data.Count fields, not offering the additional Excel properties in the .NET framework of excel accessed by Matlab if Excel is installed. 
Instead, xlread offers the process function direct access to the Excel sheet provided by the apache POI libraries under Data.WorkSheet.


