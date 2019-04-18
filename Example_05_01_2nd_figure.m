function Example_05_01_2nd_figure
% Текстовый интерпретатор - TeX

% Формируем 3D-функцию sin(x)/x
[X,Y] = meshgrid(-5:.1:5, -5:.1:5);
f = X.^2 - Y.^2;
figN = 1;
close(figure(figN)); figure(figN);
surf(X, Y, f); shading interp;

x_label_text = '\it\fontsize{18}\fontname{Times New Roman}x_{\rm\fontsize{8}1}';
y_label_text = '\it\fontsize{18}\fontname{Times New Roman}x_{\rm\fontsize{8}2}';
z_label_text = ['\fontsize{18}\fontname{Times New Roman}' ...
    '\it{f}(\itx_{\rm\fontsize{8}1}, \itx_{\rm\fontsize{8}2})'];
xlabel(x_label_text); ylabel(y_label_text); zlabel(z_label_text);

annotation_text = ['\fontname{Times New Roman}\fontsize{14}' 10 ...
    '{\itf}({\itx}) = ' ...
    '\it{\alpha} \cdot ' ...
    '({\itx}_{\rm\fontsize{8}1}^{\rm\fontsize{8}2} - ' ...
    '{\itx}_{\rm\fontsize{8}2}^{\rm\fontsize{8}2})'];
annotation('textarrow', [0 0], [1 1], ...
    'Position', [0.7482 0.8298 -0.1393 -0.3000], ...
    'HorizontalAlignment', 'center', ...
    'String', annotation_text);

pause(1);

title('\fontname{Times New Roman}\fontsize{16}\itИллюстрация возможностей интерпретатора TeX');
msword_type_text(sprintf('Верхние и нижние индексы в TeX\n'), ...
    'Заголовок 1');
str_array = {x_label_text y_label_text z_label_text annotation_text};
for i = 1:numel(str_array)
    msword_type_text([strrep(str_array{i}, '\', '\\') 10], '_Текст_');
end
msword_copy_figure(figN);