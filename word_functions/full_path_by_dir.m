function [full_path] = full_path_by_dir(dir_name)
% ѕолучение полного пути по имени директории

full_path = [];
p = path;
sep_idxs = [0 strfind(p, ';')];
for i = 1:numel(sep_idxs) - 1
    cur_path = p((sep_idxs(i) + 1):(sep_idxs(i + 1) - 1));
    if contains(cur_path, dir_name)
        full_path = cur_path;
        break;
    end
end