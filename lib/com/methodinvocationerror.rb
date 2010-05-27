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

  def initialize(message, method = "", server = "", code = -1,
                 hresult_code = -1, hresult_message = "")
    super message
    @message = message
    @method = method
    @server = server
    @code = code
    @hresult_code = hresult_code
    @hresult_message = hresult_message
  end

  def to_s
    "%s:%s:%s (0x%x)" % [@server, @method, @message, @code]
  end

  attr_reader :message, :method, :server, :code, :hresult_code, :hresult_message
end
