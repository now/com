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
    @ole.invoke(method, *args.map{ |e| e.respond_to?(:to_com) ? e.to_com : e })
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
    @ole.setproperty(property, *args.map{ |e| e.respond_to?(:to_com) ? e.to_com : e })
  rescue NoMethodError => e
    error = NoMethodError.new("undefined property `%s' for %p" % [property, self],
                              property,
                              args)
    error.set_backtrace e.backtrace.reject{ |s| s.start_with? BacktraceFilter }
    raise error
  rescue WIN32OLERuntimeError => e
    raise COM::Wrapper.raise_in('%s=' % property, e)
  end

  def load_constants(into, typelib = nil)
    saved_verbose, $VERBOSE = $VERBOSE, nil
    begin
      begin
        WIN32OLE.const_load @ole, into
      rescue RuntimeError
        begin
          @ole.ole_typelib
        rescue RuntimeError
          raise unless typelib
          WIN32OLE_TYPELIB.new(typelib)
        end.ole_types.select{ |e| e.visible? and e.ole_type == 'Enum' }.each do |enum|
          enum.constants.each do |constant|
            into.const_set constant.const_name, constant.value
          end
        end
      end
    ensure
      $VERBOSE = saved_verbose
    end
  end

  def observe(event, observer = (default = true; nil), &block)
    if default
      events.observe event, &block
    else
      events.observe event, observer, &block
    end
    self
  end

  def unobserve(event, observer = nil)
    events.unobserve event, observer
    self
  end

  def to_com
    @ole
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
