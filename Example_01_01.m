function Example_01_01
% �������� ������� Word.Application � �����������

h = h_msw;
h.Visible = true;
% get(h) - �������� (getter)
% set(h) - �������� (setter)
% invoke(h) - ������
hDoc = h.Documents.Add;
hDoc.SaveAs2(fullfile(pwd, '������ ��������'));
hDoc.SaveAs2(fullfile(pwd, '������ �������� (�����)'));
hDoc.Save;
hDoc.Close;
hDoc = h.Documents.Open(fullfile(pwd, '������ �������� (�����).docx'));

hDoc.Close;
if h.Documents.Count > 0
    h.Documents.Close;
end

AUTO_WORD_REPORT(1, '�������������� �������� ������ �� ������������ ������.');
