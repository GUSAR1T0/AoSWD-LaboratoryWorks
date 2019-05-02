function [FullFileName, FileNameExt, full_path] = preprocess_filename(FileName, varargin)
default_ext = '.docx';
if numel(varargin) > 0, default_ext = varargin{1}; end
FullFileName = fullfile(pwd, FileName);
[full_path, FileName, ext] = fileparts(FullFileName);
switch ext
    case {'.doc' '.docx'}
        default_ext = ext;
    otherwise, FileName = [FileName ext];
end
FileNameExt = [FileName default_ext];
FullFileName = fullfile(full_path, FileNameExt);
