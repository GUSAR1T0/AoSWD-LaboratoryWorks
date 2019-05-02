function [h] = h_xls
% Создание ActiveX сервера для Word Application

persistent h_actx_excel_server;
try
    if ~isempty(h_actx_excel_server)
        h_actx_excel_server.Visible = true;
    end
catch err
    if ~isempty(h_actx_excel_server)
        delete(h_actx_excel_server);
    end
    h_actx_excel_server = [];
end

if isempty(h_actx_excel_server)
    h_actx_excel_server = actxserver('Excel.Application');
end

h = h_actx_excel_server;