function figN = make_LaTeX_example(figN, Description, LaTeX_String)
msword_type_text([Description 10], 'Заголовок 1');
if ~iscell(LaTeX_String), LaTeX_String = {LaTeX_String}; end
for i = 1:numel(LaTeX_String)
    figN = figureN(figN);
    set(figN, 'Position', [1 41 1680 940]);
    text(0.5, 0.5, LaTeX_String{i}, ...
        'Interpreter', 'latex', 'FontSize', 20, 'HorizontalAlignment', 'center');
    set(gca, 'XTick', [], 'XColor', 'w', 'YTick', [], 'YColor', 'w');
    pause(1);
    msword_type_text([strrep(LaTeX_String{i}, '\', '\\') 10], '_Текст_');
    msword_copy_figure(figN);
end