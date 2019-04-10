function Example_04_01
% Вставка таблицы

hDoc = msword_open('InsertTables');
hDoc.Activate;

NameRows = {'\bfВариант 1\bf'; '\bfВариант 2\bf'};
NameCols = {'\bfN\bf' '\bf\itf\it(\itx\it)\bf', '\bf\itx\it\_1\_\bf', '\bf\itx\it\_2\_\bf',...
    '\bf\itx\it\_3\_\bf'};
TableData = {1, 2, 'нет', 4; 'да', 6, 'ошибка', 8};
TableName = 'Результат оптимизации';
msword_add_table(TableName, NameRows, NameCols, TableData);