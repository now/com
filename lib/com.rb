# -*- coding: utf-8 -*-

require 'win32ole'
WIN32OLE.codepage = WIN32OLE::CP_UTF8

# COM is an object-oriented wrapper around WIN32OLE.  COM makes it easy to add
# behavior to WIN32OLE objects, making them easier to work with from Ruby.
module COM
  autoload :Error, 'com/error'
  autoload :Events, 'com/events'
  autoload :Instantiable, 'com/instantiable'
  autoload :MethodMissing, 'com/methodmissing'
  autoload :Object, 'com/object'
  autoload :PatternError, 'com/patternerror'
  autoload :Version, 'com/version'
  autoload :Wrapper, 'com/wrapper'

  class << self
    # Gets the iconv character set equivalent of the current COM code page.
    #
    # @raise [RuntimeError] If no iconv charset is associated with the current
    #   COM codepage
    # @return [String] The iconv character set
    def charset
      COMCodePageToIconvCharset[WIN32OLE.codepage] or
        raise 'no iconv charset associated with current COM codepage: %s' %
          WIN32OLE.codepage
    end

    # Connects to a running COM object.
    #
    # This method shouldn’t be used directly, but rather is used by
    # COM::Object.
    #
    # @param [String] id The program ID of the COM object to connect to
    # @raise [COM::Error] Any error that may have occurred while trying to
    #   connect
    # @return [COM::MethodMissing] The running COM object wrapped in a
    #   COM::MethodMissing
    def connect(id)
      WIN32OLE.connect(id).extend(MethodMissing)
    rescue WIN32OLERuntimeError => e
      raise Error.from(e)
    end

    # Creates a new COM object.
    #
    # This method shouldn’t be used directly, but rather is used by
    # COM::Object.
    #
    # @param [String] id The program ID of the COM object to create
    # @param [String, nil] host The host of the program ID
    # @raise [COM::Error] Any error that may have occurred while trying to
    #   create the COM object
    # @return [COM::MethodMissing] The COM object wrapped in a COM::MethodMissing
    def new(id, host = nil)
      WIN32OLE.new(id, host).extend(MethodMissing)
    rescue WIN32OLERuntimeError => e
      raise Error.from(e)
    end
  end

private

  # @private
  COMCodePageToIconvCharset = {
    WIN32OLE::CP_UTF8 => 'UTF-8'
  }.freeze
end

require 'com/methodinvocationerror'
require 'com/pathname'
require 'com/standarderror'
require 'com/win32ole'
