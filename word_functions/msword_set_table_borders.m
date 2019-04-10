% LineWidth - {'thick' 'thin'}
function msword_set_table_borders(hTable, varargin)
[Rows, Cols, LineWidth, ...
    flTop, flLeft, flBottom, flRight] = ...
        default_args('all', 'all', 'thin', 1, 1, 1, 1, varargin{:});
if strcmp(Rows, 'all')
    Rows = 1:hTable.Rows.Count;
end
if strcmp(Cols, 'all')
    Cols = 1:hTable.Columns.Count;
end
for i = 1:numel(Rows)
    hTable.Rows.Item(i).Select;
    ApplyBorderStyle(LineWidth, ...
        flTop, flLeft, flBottom, flRight);
end
for i = 1:numel(Cols)
    hTable.Columns.Item(i).Select;
    ApplyBorderStyle(LineWidth, ...
        flTop, flLeft, flBottom, flRight);
end

function ApplyBorderStyle(LineWidth, ...
    flTop, flLeft, flBottom, flRight)
h = h_msw;
hSelection = h.Selection;
BorderTypes = {'wdBorderTop' 'wdBorderLeft' 'wdBorderBottom' 'wdBorderRight'};
flApply = boolean([flTop, flLeft, flBottom, flRight]);
IDXs = 1:numel(BorderTypes);
for i=IDXs(flApply)
    hBorder = hSelection.Borders.Item(BorderTypes{i});
    hBorder.LineStyle = 'wdLineStyleSingle';
    hBorder.LineWidth = lineWidthName(LineWidth);
    hBorder.Color = 'wdColorAutomatic';
end

function strLineWidth = lineWidthName(LineWidth)
if isnumeric(LineWidth)
    strLineWidth = sprintf('wdLineWidth%03dpt', LineWidth);
else
    switch LineWidth
        case 'thin', strLineWidth = 'wdLineWidth050pt';
        case 'thick', strLineWidth = 'wdLineWidth150pt';
    end
end