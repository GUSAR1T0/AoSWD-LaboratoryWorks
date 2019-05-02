% hWorkbook = xls_open(FileName, [mode])
% mode - 'append', 'rewrite'
function [hSheet, hWorkbook] = xls_open(WorkbookName, varargin)
mode = 'rewrite'; SheetName = [];
if ~isempty(varargin)
    SheetName = varargin{1};
    if numel(varargin) > 1, mode = varargin{2}; end
end

[h] = h_xls;
[FullFileName, FileNameExt, full_path] = preprocess_filename(WorkbookName, '.xlsx');
hWorkbook = [];
for i=1:h.Workbooks.Count
    if strcmp(h.Workbooks.Item(i).Name, FileNameExt)
        hWorkbook = h.Workbooks.Item(i);
    end
end
if isempty(hWorkbook)
    if exist(FullFileName, 'file')
        hWorkbook = h.Workbooks.Open(FullFileName);
    else
        hWorkbook = h.Workbooks.Add;
    end
end
if ~exist(full_path, 'dir')
    mkdir(full_path);
end

if ~strcmp(hWorkbook.FullName, FullFileName)
    hWorkbook.SaveAs(FullFileName);
end
hWorkbook.Save;

if isempty(SheetName)
    if hWorkbook.Sheets.Count > 0
        hSheet = hWorkbook.Sheets.Item(1);
    else
        hSheet = hWorkbook.Sheets.Add;
    end
else
    try
        hSheet = hWorkbook.Sheets.Item(SheetName);
    catch
        hWorkbook.Sheets.Add(SheetName);
    end
end

switch mode
    case 'append'
        xls_command('Ctrl+End');
    case 'rewrite'
        xls_command('Ctrl+A', 'Del', 'Ctrl+Home');
end
hWorkbook.Activate;