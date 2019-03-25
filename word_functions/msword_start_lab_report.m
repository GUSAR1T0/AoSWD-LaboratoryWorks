function hDoc = msword_start_lab_report(FileName, Lab_N, Subject, ...
    Teacher, Student, Group, Goal)
% �������� ������

[h] = h_msw;
h.Visible = true;
hDoc = h.Documents.Add;
FullFileName = fullfile(pwd, FileName);
[full_path, ~, ext] = fileparts(FullFileName);
if ~exist(full_path, 'dir')
    mkdir(full_path);
end
if ~strcmp(ext, '.docx')
    FullFileName = [FullFileName '.docx'];
end
hDoc.SaveAs2(FullFileName);
hDoc.Save;

h.Selection.Font.Name = 'Times New Roman';
h.Selection.Font.Size = 12;

h.Selection.TypeText(['M���������� ������' 13 ... 
'����������� ��������������� ���������' 13 ...
'��������������� ���������� ������� �����������' 13 ...
'������������� ��������������� �����������' 13 ...
'����������� ��. �.�.���������' 13 13 13 13 13 13 13 13 13 13 13 13 13 13 ...
]);
h.Selection.TypeParagraph;

h.Selection.WholeStory;
h.Selection.ParagraphFormat.SpaceAfter = 0;
h.Selection.HomeKey;

[word_func_path] = full_path_by_dir('word_functions');
h_pict = h.Selection.InlineShapes.AddPicture(...
    fullfile(word_func_path, 'BMP', 'AlekseevRE.png'), false, true);
h_pict.Select;
h.Selection.ShapeRange.WrapFormat.Type = 'wdWrapSquare';

h.Selection.EscapeKey;
h.Selection.MoveRight;
for i = 1:8
    h.Selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
    h.Selection.MoveDown;
end

h.Selection.EndKey;
h.Selection.MoveDown;

h.Selection.Font.Size = 14;
h.Selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
h.Selection.TypeText(['' 13 13]);
h.Selection.TypeText(['�������� ���������������� � �������������� ����������' 13 ...
'������� ����������� � ������ ����������' 13 13 13 13 13]);
h.Selection.Font.Size = 18;
h.Selection.TypeText(sprintf('\n������������ ������ � %d\n', Lab_N));
h.Selection.Font.Size = 14;
h.Selection.TypeText(sprintf('�� ����������\n'));
h.Selection.Font.Size = 16;
h.Selection.TypeText(sprintf('%s\n\n\n\n\n\n\n', Subject));
h.Selection.Font.Size = 11;
h.Selection.ParagraphFormat.LeftIndent = 226.8;
h.Selection.TypeText(sprintf('\n������������:\n________________            %s\n', Teacher));
h.Selection.Font.Size = 10;
h.Selection.ParagraphFormat.LeftIndent = 115;
h.Selection.TypeText(sprintf('  (�������)\n'));
h.Selection.Font.Size = 11;
h.Selection.ParagraphFormat.LeftIndent = 226.8;
h.Selection.TypeText(sprintf('\n�������:\n________________            %s\n', Student));
h.Selection.Font.Size = 10;
h.Selection.ParagraphFormat.LeftIndent = 115;
h.Selection.TypeText(sprintf('  (�������)\n'));
h.Selection.Font.Size = 10;
h.Selection.ParagraphFormat.LeftIndent = 226.8;
h.Selection.TypeText(sprintf('\t\t\t\t%s\n\n', Group));
h.Selection.Font.Size = 12;
h.Selection.TypeText(sprintf('������ �������� �___� _________________\n'));
h.Selection.TypeText(sprintf('� ������� _____________________________\n\n\n'));
h.Selection.ParagraphFormat.LeftIndent = 0;
h.Selection.TypeText(sprintf('������ ��������, %s', datestr(now, 'YYYY')));
h.Selection.InsertBreak(7);
h.Selection.Font.Size = 16;
h.Selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
h.Selection.TypeText(sprintf('���� ������\n'));
h.Selection.Font.Size = 12;
h.Selection.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';
h.Selection.TypeText(sprintf('%s\n', Goal));

