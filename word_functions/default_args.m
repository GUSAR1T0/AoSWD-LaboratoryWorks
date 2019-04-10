function [varargout] = default_args(varargin)
for i=1:nargout
    if numel(varargin) >= nargout+i
       varargout{i} = varargin{nargout+i};
    else
       varargout{i} = varargin{i};
    end
end