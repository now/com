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

    # @private method used by {.from}.
    def replaces?(error)
      true
    end

  protected

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
