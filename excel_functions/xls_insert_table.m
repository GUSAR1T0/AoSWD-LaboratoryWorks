function xls_insert_table(hSheet, RowCol, NameCols, NameRows, TableData)
NameRowsSize = ~isempty(NameRows);
NameColsSize = ~isempty(NameCols);

if ~isempty(NameCols)
    hRange = xls_insert(hSheet, ...
        RowCol + [0, numel(NameCols) <= size(TableData, 1)], NameCols);
    hRange.Font.FontStyle = 'полужирный';
end

if ~isempty(NameRows)
    hRange = xls_insert(hSheet, RowCol + [NameColsSize, 0], NameRows);
    hRange.Font.FontStyle = 'полужирный';
end

if ~isempty(TableData)
    hRange = xls_insert(hSheet, RowCol + [NameColsSize, NameRowsSize], TableData);
    hRange.Font.FontStyle = 'обычный';
end

RowColEnd = RowCol + [NameRowsSize NameColsSize] + size(TableData) - 1;
[hRange] = xls_Range_by_StartEnd(hSheet, RowCol, RowColEnd);
hRange.Select;
h = h_xls;
h.Selection.Rows.AutoFit;
h.Selection.Columns.AutoFit;
