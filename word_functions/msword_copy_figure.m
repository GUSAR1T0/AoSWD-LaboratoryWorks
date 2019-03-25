function msword_copy_figure(h_figure)
% Считываем заголовок графического окна

h_axes = findall(h_figure, 'Type', 'axes');
h_axis = [];
title_text = '';
for i = 1:numel(h_axes)
    title_text = get(get(h_axes(i), 'Title'), 'String');
    if ~isempty(title_text)
        h_axis = h_axes(i);
        set(get(h_axis, 'Title'), 'String', '');
        break
    end
end

tmp_file_name = 'tmp_tmp_tmp.bmp';
figure_color = get(h_figure, 'Color');
set(h_figure, 'Color', 'w');
FRAME = getframe(h_figure);
set(h_figure, 'Color', figure_color);

set(get(h_axis, 'Title'), 'String', title_text);

IMG = FRAME.cdata;
WHITE_IDXs = all(IMG == 255, 3);
IMG(:, all(WHITE_IDXs, 1), :) = [];
IMG(all(WHITE_IDXs, 2), :, :) = [];
imwrite(IMG, tmp_file_name);

h = h_msw;
hDoc = msword_active_document;
h.Selection.TypeParagraph;

h_pict = h.Selection.InlineShapes.AddPicture(...
    fullfile(pwd, tmp_file_name), false, true);
h_pict.Select;
h.Selection.ShapeRange.WrapFormat.Type = 'wdWrapInline';

h.Selection.Style = msword_get_style('_Рисунок_');
msword_command('EscapeKey', 'MoveRight', 'TypeParagraph');

msword_insert_caption('Рисунок', '_Рисунок: Подпись_', title_text);