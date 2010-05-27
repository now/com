# -*- coding: utf-8 -*-

class COM::Class
  def self.program_id(id = nil)
    @id = id if id
    return @id if @id
    raise ArgumentError,
      'No COM program ID for class: %s' % self unless
        matches = /^.*?([^:]+)::([^:]+)$/.match(name)
    @id = '%s.%s' % matches[1..2]
  end

  def self.connect
    @connect = true
  end

  def self.connect?
    @connect
  end

  def self.constants(constants)
    @constants = constants
  end

  def self.constants?
    return true unless instance_variable_defined? :@constants
    @constants
  end

  def initialize(options = {})
    connect if options.fetch(:connect, self.class.connect?)
    @object = COM.new(self.class.program_id) unless connected?
    self.class.load_constants(@object) if
      options.fetch(:constants, self.class.constants?)
  end

  def connected?
    @connected
  end

  def respond_to?(method)
    @object.ole_method(method.to_s) rescue false
  end

  # Set a bunch of properties, yield, and then restore them.  If an exception
  # is raised, any set properties are restored.
  #
  # @param [#to_hash] properties properties with values to set
  def with_properties(properties)
    saved_properties = []
    begin
      properties.to_hash.each do |property, value|
        saved_properties << [property, @object[property]]
        @object[property] = value
      end
      yield
    ensure
      previous_error = $!
      begin
        saved_properties.reverse.each do |property, value|
          begin @object[property] = value; rescue WIN32OLERuntimeError; end
        end
      rescue
        raise if not previous_error
      end
    end
  end

protected

  attr_reader :object

private

  def method_missing(*args)
    result = @object.method_missing(*args)
  rescue WIN32OLERuntimeError => e
    COM::Error.from(e)
  end

  def connect
    @object = COM.connect(self.class.program_id)
    @connected = true
  rescue COM::OperationUnavailableError
  end

  def self.load_constants(object)
    return if @constants_loaded
    modul = nesting[-2]
    saved_verbose, $VERBOSE = $VERBOSE, nil
    begin
      WIN32OLE.const_load object, modul
    ensure
      $VERBOSE = saved_verbose
    end
    @constants_loaded = true
  end
  private_class_method :load_constants

  def self.nesting
    result = []
    name.split(/::/).inject(Module) do |modul, name|
      c = modul.const_get(name) ; result << c ; c
    end
    result
  end
  private_class_method :nesting
end
