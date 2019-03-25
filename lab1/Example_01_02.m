function Example_01_02
% �������� ���������� ����� ��� ������������ ������
h = h_msw;
h.Visible = true;
hDoc = h.Documents.Add;
hDoc.SaveAs2(fullfile(pwd, '������ ����� �� ������������ ������'));
hDoc.Save;

h.Selection.Font.Name = 'Times New Roman';
h.Selection.Font.Size = 12;

h.Selection.TypeText(['����������� ������' 13 ...
'����������� ��������������� ���������' 13 ...
'��������������� ���������� ������� �����������' 13 ...
'������������� ��������������� �����������' 13 ...
'����������� ��. �.�.���������' 13]);
h.Selection.TypeParagraph

h.Selection.WholeStory;
h.Selection.ParagraphFormat.SpaceAfter = 0;
h.Selection.HomeKey;

word_function_path = full_path_by_dir('word_functions');
h_pict = h.Selection.InlineShapes.AddPicture(...
    fullfile(word_function_path, 'BMP', 'AlekseevRE.png'), false, true);
h_pict.Select;
h.Selection.ShapeRange.WrapFormat.Type = 'wdWrapSquare';

h.Selection.EscapeKey;
h.Selection.MoveRight;
for i = 1:8
    h.Selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
    h.Selection.MoveDown;
end

hDoc.Close;
if h.Documents.Count > 0
    h.Documents.Close;
end
