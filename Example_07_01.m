% �������������� Matlab � Excel
function Example_07_01
[hSheet, hWorkBook] = xls_open('������ 7. ������ � Excel');
NameRows = {'������� 1'; '������� 2'};
NameCols = {'�', 'f(x)', 'x1', 'x2', 'x3'};
TableData = {1, 2, '���', 4; '��', 6, '������', 8};
xls_insert_table(hSheet, [1 1], NameCols, NameRows, TableData);