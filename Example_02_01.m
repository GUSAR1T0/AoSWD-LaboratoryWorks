function Example_02_01
% Копирование графических окон из Matlab в MS Word

figN = 1;
figure(figN); plot(randn(50), 'o-');
Position = get(figN, 'Position');
Position([1 2]) = [50 50];
set(figN, 'Position', Position);
grid on;
format_str = '\it\fontname{Times New Roman}\fontsize{16}';
xlabel([format_str 'X']);
ylabel([format_str 'Y']);
title([format_str 'Группа случайных процессов']);

msword_copy_figure(figN);