# -*- coding: utf-8 -*-

# This module provides a wrapper around WIN32OLE’s #method_missing.  This
# wrapper deals with converting errors to the appropriate type.
#
# This is an internal module.
module COM::MethodMissing
  BacktraceFilter = File.dirname(File.dirname(__FILE__)) + File::SEPARATOR

  def method_missing(method, *args)
    super
  rescue WIN32OLERuntimeError => e
    clean = e.backtrace.reject{ |s| s.start_with? BacktraceFilter }
    raise COM::Error.from(e, clean.unshift(clean.first.sub(/:in `.*'$/, ":in `#{method}'")))
  end
  private :method_missing
end
