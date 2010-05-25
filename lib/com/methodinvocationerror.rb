# -*- coding: utf-8 -*-

class COM::MethodInvocationError < COM::Error
  extend COM::PatternError

  pattern %r{^\s*(\w+)\n
              \s*OLE\serror\scode:([0-9a-fA-F]+)
              \s+in\s+([^\n]+)\n
              \s*([^\n]+)\n
              \s*HRESULT\serror\scode:0x([0-9a-fA-F]+)[^\n]*\n
              \s*(.+)$}xu

  def self.replace(error)
    m = pattern.match(error.message)
    new(m[4], m[1], m[3], m[2].to_i(16), m[5].to_i(16), m[6])
  end

  attr_reader :message, :method, :server, :error_code, :hresult_error_code,
              :hresult_message

  def initialize(message, method = "", server = "", error_code = -1,
                 hresult_error_code = -1, hresult_message = "")
    super message
    @message = message
    @method = method
    @server = server
    @error_code = error_code
    @hresult_error_code = hresult_error_code
    @hresult_message = hresult_message
  end

  def to_s
    "%s:%s:%s (0x%x)" % [@server, @method, @message, @error_code]
  end
end
