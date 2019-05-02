function hRange = xls_insert(hSheet, RowCol, CellData)
hRange = hSheet.Range(xls_RangeName(RowCol));
TopLeft = [hRange.Row hRange.Column];
DownRight = TopLeft + (size(CellData) - 1);
RangeName = [xls_RangeName(RowCol) ':' xls_RangeName(DownRight)];
hRange = hSheet.Range(RangeName);
hRange.Value = CellData;