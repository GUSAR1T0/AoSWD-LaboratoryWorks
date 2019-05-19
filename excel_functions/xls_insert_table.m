% Merge - [Row Col RowCount ColCount]
function xls_insert_table(hSheet, RowCol, NameCols, NameRows, TableData, varargin)
Merge = [];

if ~isempty(varargin), Merge = varargin{1}; end
NameRowsSize = size(NameRows,2);
NameColsSize = size(NameCols,1);
h = h_xls;

if ~isempty(NameCols)
    hRange = xls_insert(hSheet, ...
        RowCol + [0, size(NameCols, 2) <= size(TableData, 2)], NameCols);
    hRange.Font.FontStyle = 'полужирный';
    hRange.Select;
    h.Selection.HorizontalAlignment = -4108;
    h.Selection.VerticalAlignment = -4108;
end

if ~isempty(NameRows)
    hRange = xls_insert(hSheet, RowCol + [NameColsSize, 0], NameRows);
    hRange.Font.FontStyle = 'полужирный';
    hRange.Select;
    h.Selection.HorizontalAlignment = -4108;
    h.Selection.VerticalAlignment = -4108;
end

if ~isempty(TableData)
   hRange = xls_insert(hSheet, RowCol + [NameColsSize, NameRowsSize], ...
       TableData);
   hRange.Font.FontStyle = 'обычный';
   hRange.Select;
   h.Selection.HorizontalAlignment = -4152;
   h.Selection.VerticalAlignment = -4108;
end

for i=1:size(Merge, 1)
    [hRange] = xls_Range_by_StartEnd(hSheet, Merge(i, 1:2), ...
        Merge(i, 1:2) + (Merge(i, 3:4) - 1));
    hRange.Select;
    h.Selection.Merge;
end

RowColEnd = RowCol + [NameColsSize NameRowsSize] + size(TableData) - 1;
[hRange] = xls_Range_by_StartEnd(hSheet, RowCol, RowColEnd);
hRange.Select;
h.Selection.Rows.AutoFit;
h.Selection.Columns.AutoFit;
ColWidth = zeros(1, RowColEnd(2));

for i = (RowCol(2) + NameRowsSize):RowColEnd(2)
    ColWidth(i) = h.Column.Item(i).ColumnWidth;
end

for i = (RowCol(2) + NameRowsSize):RowColEnd(2)
    h.Column.Item(i).ColumnWidth = max(ColWidth);
end

for i = 1:4
    h.Selection.Borders.Item(i).LineStyle=1;
end
