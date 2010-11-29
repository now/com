# -*- coding: utf-8 -*-

class COM::Error < RuntimeError
  class << self
    def from(error, backtrace = nil)
      errors.find{ |replacement| replacement.replaces? error }.replace(error).tap{ |e|
        e.set_backtrace backtrace if backtrace
      }
    end

    def replaces?(error)
      true
    end

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
