function msword_type_text(str, varargin)
style = '_Текст_';
if numel(varargin) > 0
    style = varargin{1};
end
h = h_msw;
h.Selection.Style = msword_get_style(style);
tokens = msword_tokens;
[cell_str, token_idxs] = split_string_by_tokens(str, tokens(:,1)');
for i=1:(numel(cell_str))
    switch cell_str{i}
        case tokens(:,1)'
            msword_command(tokens{token_idxs(i), 2});
        otherwise
            h.Selection.TypeText(cell_str{i});
    end
end
