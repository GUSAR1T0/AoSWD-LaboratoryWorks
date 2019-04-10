% msword_command(cmd1, cmd2, ...)
% cmd1, cmd2, ... - commands
function msword_command(varargin)
h = h_msw;
for i=1:numel(varargin)
    switch upper(varargin{i})
        case upper({'Ctrl+A'})
            h.Selection.WholeStory;
        case upper({'Ctrl+B'})
            switch h.Selection.Font.Bold
                case -1, h.Selection.Font.Bold = 0;
                case 0, h.Selection.Font.Bold = -1;
            end
        case upper({'Ctrl+End'})
            h.Selection.EndKey(6); %6 - end of document, 5 - end of string
        case upper({'Ctrl+I'})
            switch h.Selection.Font.Italic
                case -1, h.Selection.Font.Italic = 0;
                case 0, h.Selection.Font.Italic = -1;
            end
        case upper({'Ctrl+Ins'})
            h.Selection.InsertSymbol(-4051, 'Symbol', true);
        case upper({'Ctrl+Shift++'})
            switch h.Selection.Font.Superscript
                case -1, h.Selection.Font.Superscript = 0;
                case 0, h.Selection.Font.Superscript = -1;
            end
        case upper({'Ctrl+U'})
            switch h.Selection.Font.Underline
                case 'wdUnderlineSingle', h.Selection.Font.Underline = 'wdUnderlineNone';
                case 'wdUnderlineNone', h.Selection.Font.Underline = 'wdUnderlineSingle';
            end
        case upper({'Ctrl++'})
            switch h.Selection.Font.Subscript
                case -1, h.Selection.Font.Subscript = 0;
                case 0, h.Selection.Font.Subscript = -1;
            end
        case upper({'Del' 'Delete'})
            h.Selection.Delete;
        case upper({'Enter' 'TypeParagraph'})
            h.Selection.TypeParagraph;
        case upper({'Esc' 'EscapeKey'})
            h.Selection.EscapeKey;
        case upper({'Left' 'MoveLeft' '<-'})
            h.Selection.MoveLeft;
        case upper({'Right' 'MoveRight' '->'})
            h.Selection.MoveRight;    
    end
end