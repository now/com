# -*- coding: utf-8 -*-

class COM::Instantiable < COM::Object
  class << self
    def program_id(id = nil)
      @id = id if id
      return @id if instance_variable_defined? :@id
      raise ArgumentError,
        'no COM program ID for class: %s' % self unless
          matches = /^.*?([^:]+)::([^:]+)$/.match(name)
      @id = '%s.%s' % matches[1..2]
    end

    def connect
      @connect = true
    end

    def connect?
      @connect ||= false
    end

    def constants(constants)
      @constants = constants
    end

    def constants?
      return true unless instance_variable_defined? :@constants
      @constants
    end

    def load_constants(com)
      return if constants_loaded?
      modul = nesting[-2]
      saved_verbose, $VERBOSE = $VERBOSE, nil
      begin
        WIN32OLE.const_load com, modul
      ensure
        $VERBOSE = saved_verbose
      end
      @constants_loaded = true
    end

    def constants_loaded?
      @constants_loaded ||= false
    end

    def nesting
      result = []
      name.split(/::/).inject(Module) do |modul, name|
        modul.const_get(name).tap{ |c| result << c }
      end
      result
    end
  end

  def initialize(options = {})
    @connected = false
    connect if options.fetch(:connect, self.class.connect?)
    self.com = COM.new(self.class.program_id) unless connected?
    self.class.load_constants(com) if
      options.fetch(:constants, self.class.constants?)
  end

  def connected?
    @connected
  end

private

  def connect
    self.com = COM.connect(self.class.program_id)
    @connected = true
  rescue COM::OperationUnavailableError
  end
end
