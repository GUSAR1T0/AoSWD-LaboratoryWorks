% xls_command(cmd1, cmd2, ...)
% cmd1, cmd2, ... - commands
function xls_command(varargin)
h = h_xls;
for i=1:numel(varargin)
    switch upper(varargin{i})
        case upper({'Ctrl+A'})
            h.Cells.Select;
        case upper({'Del'})
            h.Selection.ClearContents;
        case upper({'Ctrl+End'})
            h.ActiveCell.SpecialCells(xls_cell_type('xlLastCell')).Select;
        case upper({'Ctrl+Home'})
            h.Range('A1').Select;
    end
end