# -*- coding: utf-8 -*-

class COM::Class
  def initialize(com)
    @com = com
  end

  def respond_to?(method)
    com.ole_method(method.to_s) rescue false
  end

  # Set a bunch of properties, yield, and then restore them.  If an exception
  # is raised, any set properties are restored.
  #
  # @param [#to_hash] properties properties with values to set
  def with_properties(properties)
    saved_properties = []
    begin
      properties.to_hash.each do |property, value|
        saved_properties << [property, com[property]]
        com[property] = value
      end
      yield
    ensure
      previous_error = $!
      begin
        saved_properties.reverse.each do |property, value|
          begin com[property] = value; rescue WIN32OLERuntimeError; end
        end
      rescue
        raise if not previous_error
      end
    end
  end

protected

  attr_reader :com

private

  def method_missing(*args)
    com.method_missing(*args)
  rescue WIN32OLERuntimeError => e
    raise COM::Error.from(e)
  end
end
