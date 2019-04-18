function Example_05_01_1st_figure
% Текстовый интерпретатор - TeX

% Формируем 3D-функцию sin(x)/x
[X,Y] = meshgrid(-5:.1:5, -5:.1:5);
f = sinc(sqrt(X.^2 + Y.^2));
figN = 1;
close(figure(figN)); figure(figN);
surf(X, Y, f); shading interp;

x_label_text = '\it\fontsize{18}\fontname{Times New Roman}X';
y_label_text = '\it\fontsize{18}\fontname{Times New Roman}Y';
z_label_text = '\it\fontsize{18}\fontname{Times New Roman}Z';
xlabel(x_label_text); ylabel(y_label_text); zlabel(z_label_text);

annotation_text = ['\fontname{Times New Roman}\fontsize{12}' ...
    'Функция' 10 '{\itf}({\itx}) = sin({\itx})/{\itx}'];
annotation('textarrow', [0 0], [1 1], ...
    'Position', [0.7982 0.8583 -0.2571 -0.1476], ...
    'HorizontalAlignment', 'center', ...
    'String', annotation_text);

pause(1);

title('\fontname{Times New Roman}\fontsize{16}\itИллюстрация возможностей интерпретатора TeX');
msword_type_text(sprintf('Смена шрифтов, размеров и начертания в TeX\n'), ...
    'Заголовок 1');
str_array = {x_label_text y_label_text z_label_text annotation_text};
for i = 1:numel(str_array)
    msword_type_text([strrep(str_array{i}, '\', '\\') 10], '_Текст_');
end
msword_copy_figure(figN);