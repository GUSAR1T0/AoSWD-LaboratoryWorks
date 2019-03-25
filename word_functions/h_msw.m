function [h] = h_msw
% Создание ActiveX сервера для Word Application

persistent h_actx_word_server;
try
    if ~isempty(h_actx_word_server)
        h_actx_word_server.Visible = true;
    end
catch err
    if ~isempty(h_actx_word_server)
        delete(h_actx_word_server);
    end
    h_actx_word_server = [];
end

if isempty(h_actx_word_server)
    h_actx_word_server = actxserver('Word.Application');
end

h = h_actx_word_server;