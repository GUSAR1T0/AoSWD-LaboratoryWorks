%mode - 'append', 'rewrite'
function hDoc = msword_open(FileName, varargin)
mode = 'rewrite';
if ~isempty(varargin), mode = varargin{1}; end
[h] = h_msw;
[FullFileName, FileNameExt, full_path] = preprocess_filename(FileName);
hDoc = [];
for i=1:h.Documents.Count
    if strcmp(h.Documents.Item(i).Name, FileNameExt)
        hDoc = h.Documents.Item(i);
    end
end
if isempty(hDoc)
    hDoc = h.Documents.Add;
end
if ~exist(full_path, 'dir')
    mkdir(full_path);
end

hDoc.SaveAs2(FullFileName);
hDoc.Save;
switch mode
    case 'append'
        msword_command('Ctrl+End', 'Enter');
    case 'rewrite'
        msword_command('Ctrl+A', 'Del');
end
hDoc.Activate;