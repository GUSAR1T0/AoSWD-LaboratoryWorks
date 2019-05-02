function [hRange] = xls_Range_by_StartEnd(hSheet, RowColStart, RowColEnd)
RangeName = [xls_RangeName(RowColStart) ':' xls_RangeName(RowColEnd)];
hRange = hSheet.Range(RangeName);