# -*- coding: utf-8 -*-

class COM::Object
  # Creates a new instance based on _com_.  It is important that subclasses
  # call `super` if they override this method.
  #
  # @param [WIN32OLE] com A WIN32OLE Com object
  def initialize(com)
    self.com = com
  end

  # Queries whether this COM object responds to _method_.
  #
  # @param [Symbol] method Method name to query for response
  # @return [Boolean] Whether or not this COM object responds to _method_
  def respond_to?(method)
    super or com.respond_to? method
  end

  # Sets a bunch of properties, yield, and then restore them.  If an exception
  # is raised, any set properties are restored.
  #
  # @param [#to_hash] properties properties with values to set
  # @return [COM::Object] `self` */
  def with_properties(properties)
    saved_properties = []
    begin
      properties.to_hash.each do |property, value|
        saved_properties << [property, com[property]]
        com.set_property property, value
      end
      yield
    ensure
      previous_error = $!
      begin
        saved_properties.reverse.each do |property, value|
          begin com.set_property property, value; rescue COM::Error; end
        end
      rescue
        raise if not previous_error
      end
    end
  end

  def observe(event, observer = (default = true; nil), &block)
    if default
      com.observe event, &block
    else
      com.observe event, observer, &block
    end
    self
  end

  def unobserve(event, observer = nil)
    com.unobserve event, observer
    self
  end

  def to_com
    com.to_com
  end

  protected

  attr_reader :com

  def com=(com)
    @com = WIN32OLE === com ? COM::Wrapper.new(com) : com
  end

  private

  def method_missing(method, *args)
    case method.to_s
    when /=\z/
      com.set_property($`.encode(COM.charset), *args)
    else
      com.invoke(method.to_s.encode(COM.charset), *args)
    end
  rescue NoMethodError => e
    raise e, "undefined method `%s' for %p" % [method, self], e.backtrace
  end
end
