% �������������� Matlab � Excel
function Example_07_01
[hSheet, hWorkBook] = xls_open('������ 7. ������ � Excel');
NameRows = {'������� 1'; '������� 2'};
NameCols = {'N' 'f(x)', 'X', '', ''; '', '', 'x1', 'x2', 'x3'};
TableData = {1, 2, '���', 4; '��', 6, '�����', 8};
Merge = [1 1 2 1; 1 2 2 1; 1 3 1 3];
xls_insert_table(hSheet, [1 1], NameCols, NameRows, TableData, Merge);