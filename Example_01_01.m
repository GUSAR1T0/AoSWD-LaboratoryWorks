function Example_01_01
% Создание сервера Word.Application с автоотчетом

h = h_msw;
h.Visible = true;
% get(h) - свойства (getter)
% set(h) - свойства (setter)
% invoke(h) - методы
hDoc = h.Documents.Add;
hDoc.SaveAs2(fullfile(pwd, 'Первый документ'));
hDoc.SaveAs2(fullfile(pwd, 'Первый документ (копия)'));
hDoc.Save;
hDoc.Close;
hDoc = h.Documents.Open(fullfile(pwd, 'Первый документ (копия).docx'));

hDoc.Close;
if h.Documents.Count > 0
    h.Documents.Close;
end

AUTO_WORD_REPORT(1, 'Автоматическое создание отчета по лабораторной работе.');
