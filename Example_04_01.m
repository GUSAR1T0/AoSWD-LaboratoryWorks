function Example_04_01
% ������� �������

hDoc = msword_open('InsertTables');
hDoc.Activate;

NameRows = {'\bf������� 1\bf'; '\bf������� 2\bf'};
NameCols = {'\bfN\bf' '\bf\itf\it(\itx\it)\bf', '\bf\itx\it\_1\_\bf', '\bf\itx\it\_2\_\bf',...
    '\bf\itx\it\_3\_\bf'};
TableData = {1, 2, '���', 4; '��', 6, '������', 8};
TableName = '��������� �����������';
msword_add_table(TableName, NameRows, NameCols, TableData);