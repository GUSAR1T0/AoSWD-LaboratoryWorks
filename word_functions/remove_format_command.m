function str = remove_format_command(str)
str = strrep(str, '\it', '');
str = strrep(str, '\bf', '');
str = regexprep(str, '\\fontname{[\w\s]+}', '');
str = regexprep(str, '\\fontsize{[\d]+}', '');