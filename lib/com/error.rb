# -*- coding: utf-8 -*-

# Superclass for COM-related errors.  This class is meant to be subclassed an
# used for refinement of otherwise very unspecific COM errors.
#
# @see COM::StandardError
class COM::Error < RuntimeError
  class << self
    # Creates a COM::Error from _error_, with optional _backtrace_.  _Error_
    # should be an error raised by WIN32OLE.
    #
    # This is an internal method.
    #
    # @param [WIN32OLERuntimeError] error The error to create a new one from
    # @param [Array<String>, nil] backtrace The backtrace to use for the new
    #   error
    def from(error, backtrace = nil)
      errors.find{ |replacement| replacement.replaces? error }.replace(error).tap{ |e|
        e.set_backtrace backtrace if backtrace
      }
    end

    # Returns true if this class replaces _error_.  A class is meant to replace
    # an error that it better represents.
    #
    # @param [WIN32OLERuntimeError] error The error to check against
    # @return [Boolean] True if this class replaces _error_
    def replaces?(error)
      true
    end

    # Replaces _error_ with a more suitable one.  The {.replaces?}
    # method must have been called before this method and must have returned
    # true before this method may be called.
    #
    # @param [WIN32OLERuntimeError] error The error to replace
    # @return [COM::Error] A more suitable error than _error_
    def replace(error)
      new(error.message)
    end

  private

    def inherited(error)
      errors.unshift error
    end

    def errors
      @@errors ||= [self]
    end
  end
end
