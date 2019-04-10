function [FullFileName, FileNameExt, full_path] = preprocess_filename(FileName)
FullFileName = fullfile(pwd, FileName);
[full_path, FileName, ext] = fileparts(FullFileName);
switch ext
    case {'.doc' '.docx'}
    otherwise, FileName = [FileName ext];
end
FileNameExt = [FileName '.docx'];
FullFileName = fullfile(full_path, FileName);
