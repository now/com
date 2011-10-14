# -*- coding: utf-8 -*-

# This module provides a wrapper around WIN32OLE’s instance methods.  This
# wrapper deals with converting errors to the appropriate type.
#
# @private
class COM::Wrapper
  BacktraceFilter = File.dirname(File.dirname(__FILE__)) + File::SEPARATOR

  class << self
    def raise_in(method, e)
      clean = e.backtrace.reject{ |s| s.start_with? BacktraceFilter }
      raise COM::Error.from(e, clean.unshift(clean.first.sub(/:in `.*'$/, ":in `#{method}'")))
    end
  end

  def initialize(ole)
    @ole = ole
  end

  def invoke(method, *args)
    @ole.invoke(method, *args)
  rescue NoMethodError => e
    error = NoMethodError.new("undefined method `%s' for %p" % [method, self],
                              method,
                              args)
    error.set_backtrace e.backtrace.reject{ |s| s.start_with? BacktraceFilter }
    raise error
  rescue WIN32OLERuntimeError => e
    raise COM::Wrapper.raise_in(method, e)
  end

  def set_property(property, *args)
    @ole.setproperty(property, *args)
  rescue NoMethodError => e
    error = NoMethodError.new("undefined property `%s' for %p" % [property, self],
                              property,
                              args)
    error.set_backtrace e.backtrace.reject{ |s| s.start_with? BacktraceFilter }
    raise error
  rescue WIN32OLERuntimeError => e
    raise COM::Wrapper.raise_in('%s=' % property, e)
  end

  def load_constants(into)
    saved_verbose, $VERBOSE = $VERBOSE, nil
    begin
      begin
        WIN32OLE.const_load @ole, into
      rescue RuntimeError
        WIN32OLE_TYPE.enums(program_id).each do |enum|
          enum.constants.each do |constant|
            into.const_set constant.const_name, constant.value
          end
        end
      end
    ensure
      $VERBOSE = saved_verbose
    end
  end

  def observe(event, observer = COM::Events::ArgumentMissing, &block)
    events.observe event, observer, &block
    self
  end

  def unobserve(event, observer = nil)
    events.unobserve event, observer
    self
  end

  protected

  def events
    @events ||= COM::Events.new(@ole)
  end

  private

  def method_missing(method, *args)
    case method.to_s
    when /=\z/
      begin
        set_property($`.encode(COM.charset), *args)
      rescue NoMethodError => e
        raise e, "undefined method `%s' for %p" % [method, self], e.backtrace
      end
    else
      invoke(method.to_s.encode(COM.charset), *args)
    end
  end
end
