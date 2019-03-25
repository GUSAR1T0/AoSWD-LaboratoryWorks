function AUTO_WORD_REPORT (Lab_N, Goal)
% �������� ������ � MS Word � ����������� �����

MyName = '��������� �.�.';
Group = '�18-���-1';

hDoc = msword_start_lab_report(...
    sprintf('\\������\\%s_%s ����� �� ������������ ������ N%d', Group, MyName, Lab_N), ...
    Lab_N, '������������� ���������������� ����������� ������� ������������', ...
    '�������� �.�.', MyName, Group, ...
    Goal);

[ST, ~] = dbstack('-completenames');
if numel (ST)>1
    fid = fopen(ST(2).file, 'r');
    fun_txt = native2unicode(fread(fid)','Windows-1251');
    fclose(fid);
    h = h_msw;
    h.Selection.Font.Size = 16;
    h.Selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
    h.Selection.TypeText(sprintf('����������� ���\n'));
    h.Selection.Font.Size = 12;
    h.Selection.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';
    h.Selection.TypeText(sprintf('%s\n', fun_txt));
end

hDoc.Close;