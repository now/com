# -*- coding: utf-8 -*-

# Represents an COM method invocation error.
class COM::MethodInvocationError < COM::Error
  extend COM::PatternError

  pattern %r{^\s*([^\n]*)\n
              \s*OLE\serror\scode:([0-9a-fA-F]+)
              \s+in\s+([^\n]+)\n
              \s*([^\n]+)\n
              \s*HRESULT\serror\scode:0x([0-9a-fA-F]+)[^\n]*\n
              \s*(.+)$}xu

  class << self
    # This is an internal method used by COM::Error.
    def replace(error)
      m = pattern.match(error.message)
      new(m[4], m[1], m[3], m[2].to_i(16), m[5].to_i(16), m[6])
    end
  end

  # Creates a new COM::MethodInvocationError with _message_.
  #
  # @param [#to_str] message Error message
  # @param [String] method Method error occurred in
  # @param [String] server Server error occurred in
  # @param [Integer] code COM error code
  # @param [Integer] hresult_code HRESULT error code
  # @param [String] hresult_message HRESULT error message
  def initialize(message, method = '', server = '', code = -1,
                 hresult_code = -1, hresult_message = '')
    super message
    @method, @server, @code, @hresult_code, @hresult_message =
      method, server, code, hresult_code, hresult_message
  end

  attr_reader :method, :server, :code, :hresult_code, :hresult_message
end
