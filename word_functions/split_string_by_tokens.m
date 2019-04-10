%tokens - {'\it' '\bf'}
function [cell_str, token_idxs] = split_string_by_tokens(str, tokens)
cell_str = {}; token_idxs = [];
slash_idxs = [1 find(str=='\') length(str)+1];
dbl_slash_idxs = findstr(str, '\\');
slash_idxs = setdiff(slash_idxs, [dbl_slash_idxs dbl_slash_idxs+1]);
if slash_idxs(1) ~=1 , slash_idxs = [1 slash_idxs]; end
for i =1: (numel(slash_idxs)-1)
    cur_str = str(slash_idxs(i):(slash_idxs(i+1)-1));
    if isempty(cur_str), continue; end
    for j =1: numel(tokens)
        if strcmp(cur_str(1:min(end, length(tokens{j}))), tokens{j})
            cell_str(end+1) = tokens(j);
            token_idxs(end+1) = j;
            cur_str(1:length(tokens{j})) = [];
            break;
        end
    end
    cell_str(end+1) = {cur_str};
    token_idxs(end+1) = 0;
end