function hDoc = msword_active_document
h = h_msw;
if h.Documents.Count == 0
    h.Documents.Add;
end
hDoc = h.ActiveDocument;