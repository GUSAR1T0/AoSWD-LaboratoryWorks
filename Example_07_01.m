% Взаимодействие Matlab и Excel
function Example_07_01
[hSheet, hWorkBook] = xls_open('Лекция 7. Работа с Excel');
NameRows = {'Вариант 1'; 'Вариант 2'};
NameCols = {'N' 'f(x)', 'X', '', ''; '', '', 'x1', 'x2', 'x3'};
TableData = {1, 2, 'нет', 4; 'да', 6, 'отказ', 8};
Merge = [1 1 2 1; 1 2 2 1; 1 3 1 3];
xls_insert_table(hSheet, [1 1], NameCols, NameRows, TableData, Merge);