function Name = xls_RangeName(RowCol)
Col = RowCol(2);
Row = RowCol(1);
Digits(1) = mod(Col, 26);
Col = (Col - Digits(1)) / 26;
while Col > 0
    Digits(end + 1) = mod(Col, 26);
    Col = (Col - Digits(end)) / 26;
end
ColName = char('A' + (Digits(end:-1:1) - 1));
Name = [ColName num2str(Row)];