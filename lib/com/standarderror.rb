# -*- coding: utf-8 -*-

# Sets up mappings between COM errors and Ruby errors.
#
# @private
module COM::StandardError
  class << self
  private
    def define(code, errorclass, message = nil, &block)
      block = proc{ |m| m[1] } unless message or block
      Class.new(COM::Error) do
        extend COM::PatternError

        pattern %r{^\s*(.*)\n\s*HRESULT\serror\scode:(0x(?i:#{code.to_s(16)}))}xu

        (class << self; self; end).class_eval do
          define_method :replace do |error|
            m = pattern.match(error.message)
            errorclass.new(message ? message : block.call(m))
          end
        end
      end
    end
  end

  define(0x80004001, ::NotImplementedError){ |m| '%s: not implemented' % m[1] }
  define 0x80020005, ::TypeError
  define 0x80020006, ::NoMethodError
  define 0x8002000e, ::ArgumentError, 'wrong number of arguments'
  define 0x800401e4, ::ArgumentError
end

# Sets up mappings between HRESULT errors and COM errors.
#
# @private
module COM::HResultError
  class << self
    def define(code, error, message = nil, &block)
      block = proc{ |m| m[1] } unless message or block
      COM.const_set error, Class.new(COM::Error){
        extend COM::PatternError

        pattern %r{^\s*(.*)\n\s*HRESULT\serror\scode:(0x(?i:#{code.to_s(16)}))}xu

        (class << self; self; end).class_eval do
          define_method :replace do |error|
            m = pattern.match(error.message)
            new(message ? message : block.call(m))
          end
        end
      }
    end
  end

  define 0x80020003, :MemberNotFoundError
  define 0x800401e3, :OperationUnavailableError
end
