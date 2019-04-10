function msword_add_table(TableName, NameRows, NameCols, TableData)
h = h_msw;
hDoc = h.ActiveDocument;
FullTableData = TableData2Str([NameCols ; [NameRows TableData]]);
[NumRows, NumColumns] = size(FullTableData);
msword_insert_caption('Таблица', '_Таблица: Haзвание_', TableName);
hTable = hDoc.Tables.Add(h.Selection.Range, NumRows, NumColumns);
for i=1:NumRows
    for j=1:NumColumns
        hTable.Cell(i, j).Select;
        msword_command('<-');
        msword_type_text(FullTableData{i, j}, '_Необновляемый_');
    end
end
msword_set_table_borders(hTable, 'all', 'all', 'thin');
msword_set_table_borders(hTable, 1, 1, 'thick');

function [TableData] = TableData2Str(TableData)
for i=1:size(TableData, 1)
    for j=1:size(TableData, 2)
        if isnumeric(TableData{i, j})
            TableData{i, j} = num2str(TableData{i,j});
        end
    end
end