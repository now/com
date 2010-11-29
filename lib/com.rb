# -*- coding: utf-8 -*-

require 'win32ole'
WIN32OLE.codepage = WIN32OLE::CP_UTF8

module COM
  autoload :Error, 'com/error'
  autoload :Events, 'com/events'
  autoload :Instantiable, 'com/instantiable'
  autoload :MethodMissing, 'com/methodmissing'
  autoload :Object, 'com/object'
  autoload :PatternError, 'com/patternerror'
  autoload :Version, 'com/version'
  autoload :Wrapper, 'com/wrapper'

  COMCodePageToIconvCharset = {
    WIN32OLE::CP_UTF8 => 'UTF-8'
  }.freeze

  class << self
    def charset
      COMCodePageToIconvCharset[WIN32OLE.codepage] or
        raise 'no iconv charset associated with current COM codepage: %s' %
          WIN32OLE.codepage
    end

    def connect(id)
      WIN32OLE.connect(id).extend MethodMissing
    rescue WIN32OLERuntimeError => e
      raise Error.from(e)
    end

    def new(id, host = nil)
      WIN32OLE.new(id, host).extend MethodMissing
    rescue WIN32OLERuntimeError => e
      raise Error.from(e)
    end
  end
end

require 'com/methodinvocationerror'
require 'com/pathname'
require 'com/standarderror'
