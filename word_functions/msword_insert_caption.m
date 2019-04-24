function msword_insert_caption(CaptionName, CaptionStyle, CaptionText)
h = h_msw;
hCaptionLabel = h.CaptionLabels.Add(CaptionName);
hCaptionLabel.IncludeChapterNumber = false;
h.Selection.InsertCaption(CaptionName);
h.Selection.Style = msword_get_style(CaptionStyle);
if ~isempty(strtrim(CaptionText))
    h.Selection.TypeText(' ');
    h.Selection.InsertSymbol(-4051, 'Symbol', true);
    CaptionText = remove_format_command(CaptionText);
end
h.Selection.TypeText(sprintf(' %s', CaptionText));
h.Selection.TypeParagraph;