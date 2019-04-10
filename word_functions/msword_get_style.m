function h_style = msword_get_style(style_name)
h = h_msw;
h_style = [];
for i = 1:h.ActiveDocument.Styles.Count
    if strcmp(style_name, h.ActiveDocument.Styles.Item(i).Name)
        h_style = h.ActiveDocument.Styles.Item(i);
    end
end

if isempty(h_style)
    h_style = h.ActiveDocument.Styles.Add(style_name);
    h_style.Font.Name = 'Times New Roman';
    h_style.AutomaticallyUpdate = true;
    switch style_name
        case '_Необновляемый_'
            h_style.Font.Size = 12;
            h_style.Font.Italic = 0;
            h_style.Font.ColorIndex = 'wdAuto';
            h_style.ParagraphFormat.LeftIndent = 0;
            h_style.ParagraphFormat.RightIndent = 0;
            h_style.ParagraphFormat.SpaceBefore = 0;
            h_style.ParagraphFormat.SpaceAfter = 0;
            h_style.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';
            h_style.ParagraphFormat.KeepWithNext = 0;
            h_style.AutomaticallyUpdate = false;
        case '_Рисунок_'
            h_style.ParagraphFormat.LeftIndent = 0;
            h_style.ParagraphFormat.RightIndent = 0;
            h_style.ParagraphFormat.SpaceBefore = 10;
            h_style.ParagraphFormat.SpaceAfter = 0;
            h_style.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
            h_style.ParagraphFormat.KeepWithNext = -1;
        case '_Рисунок: Подпись_'
            h_style.Font.Size = 11;
            h_style.Font.Italic = 0;
            h_style.Font.ColorIndex = 'wdAuto';
            h_style.ParagraphFormat.LeftIndent = 0;
            h_style.ParagraphFormat.RightIndent = 0;
            h_style.ParagraphFormat.SpaceBefore = 4;
            h_style.ParagraphFormat.SpaceAfter = 10;
            h_style.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';
            h_style.ParagraphFormat.KeepWithNext = 0;
        case '_Текст_'
            h_style.Font.Size = 12;
            h_style.Font.Italic = 0;
            h_style.Font.ColorIndex = 'wdAuto';
            h_style.ParagraphFormat.LeftIndent = 0;
            h_style.ParagraphFormat.RightIndent = 0;
            h_style.ParagraphFormat.SpaceBefore = 0;
            h_style.ParagraphFormat.SpaceAfter = 0;
            h_style.ParagraphFormat.Alignment = 'wdAlignParagraphJustify';
            h_style.ParagraphFormat.KeepWithNext = 0;
    end
end