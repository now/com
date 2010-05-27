# -*- coding: utf-8 -*-

require 'facets/kernel/silence'

require 'win32ole'
WIN32OLE.codepage = WIN32OLE::CP_UTF8

module COM
  autoload :Class, 'com/class'
  autoload :Error, 'com/error'
  autoload :Events, 'com/events'
  autoload :InstantiableClass, 'com/instantiableclass'
  autoload :PatternError, 'com/patternerror'
  autoload :Version, 'com/version'
  autoload :Wrapper, 'com/wrapper'

  COMCodePageToIconvCharset = {
    WIN32OLE::CP_UTF8 => 'UTF-8'
  }.freeze

  def self.charset
    COMCodePageToIconvCharset[WIN32OLE.codepage] or
      raise 'No iconv charset associated with current COM codepage: ' %
        WIN32OLE.codepage
  end

  def self.connect(id)
    WIN32OLE.connect(id)
  rescue WIN32OLERuntimeError => e
    COM::Error.from(e)
  end

  def self.new(id, host = nil)
    WIN32OLE.new(id, host)
  rescue WIN32OLERuntimeError => e
    COM::Error.from(e)
  end
end

require 'com/methodinvocationerror'
require 'com/pathname'
require 'com/standarderror'
