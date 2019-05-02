% Взаимодействие Matlab и Excel
function Example_07_01
[hSheet, hWorkBook] = xls_open('Лекция 7. Работа с Excel');
NameRows = {'Вариант 1'; 'Вариант 2'};
NameCols = {'№', 'f(x)', 'x1', 'x2', 'x3'};
TableData = {1, 2, 'нет', 4; 'да', 6, 'ошибка', 8};
xls_insert_table(hSheet, [1 1], NameCols, NameRows, TableData);