# -*- coding: utf-8 -*-

class COM::InstantiableClass < COM::Class
  def self.program_id(id = nil)
    @id = id if id
    return @id if instance_variable_defined? :@id
    raise ArgumentError,
      'No COM program ID for class: %s' % self unless
        matches = /^.*?([^:]+)::([^:]+)$/.match(name)
    @id = '%s.%s' % matches[1..2]
  end

  def self.connect
    @connect = true
  end

  def self.connect?
    @connect ||= false
  end

  def self.constants(constants)
    @constants = constants
  end

  def self.constants?
    return true unless instance_variable_defined? :@constants
    @constants
  end

  def initialize(options = {})
    @connected = false
    connect if options.fetch(:connect, self.class.connect?)
    @object = COM.new(self.class.program_id) unless connected?
    self.class.load_constants(@object) if
      options.fetch(:constants, self.class.constants?)
  end

  def connected?
    @connected
  end

private

  def connect
    @object = COM.connect(self.class.program_id)
    @connected = true
  rescue COM::OperationUnavailableError
  end

  def self.load_constants(object)
    return if constants_loaded?
    modul = nesting[-2]
    saved_verbose, $VERBOSE = $VERBOSE, nil
    begin
      WIN32OLE.const_load object, modul
    ensure
      $VERBOSE = saved_verbose
    end
    @constants_loaded = true
  end

  def self.constants_loaded?
    @constants_loaded ||= false
  end

  def self.nesting
    result = []
    name.split(/::/).inject(Module) do |modul, name|
      c = modul.const_get(name) ; result << c ; c
    end
    result
  end
  private_class_method :nesting
end
