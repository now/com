# -*- coding: utf-8 -*-

# Represents instantiable COM objects.  These are COM objects that we can
# connect to and create.
class COM::Instantiable < COM::Object
  class << self
    # Gets or sets the COM program ID.
    #
    # If no program ID has explicitly been set, one based on the name of this
    # class and its containing module.  For example, A::B is turned into
    # 'A.B'.
    #
    # @param [String,nil] id Program ID to, if given, use
    # @return [String] The set or automatically generated program ID
    # @raise [ArgumentError] If no program ID has been set and one can’t be
    #   automatically generated
    def program_id(id = nil)
      @id = id if id
      return @id if defined? @id
      raise ArgumentError,
        'no automatic COM program ID for class available: %s' % self unless
          matches = /^.*?([^:]+)::([^:]+)$/.match(name)
      @id = '%s.%s' % matches[1..2]
    end

    # Marks this class as trying to connect to already running instances of COM
    # objects.
    #
    # @return true
    def connect
      @connect = true
    end

    # Queries whether or not this class tries to connect to already running
    # instances of COM objects.
    #
    # The default is false.
    #
    # @return Whether or not this class tries to connect to already running
    #   instances of COM objects
    #
    # @see #connect
    def connect?
      @connect ||= false
    end

    # Sets whether or not this class tries to load constants when connecting to
    # or creating a COM object.
    #
    # @param [Boolean] constants Whether or not to load constans
    # @return [Boolean] Whether or not to load constants
    def constants(constants)
      @constants = constants
    end

    # Queries whether or not this class tries to load constans when connecting
    # to or creating a COM object.
    #
    # The default is true.
    #
    # @return Whether or not to load constants
    def constants?
      return true unless defined? @constants
      @constants
    end

    # Typelib to use for loading constants, if it can’t be determined
    # automatically.
    def typelib(typelib = nil)
      @typelib = typelib if typelib
      @typelib ||= nil
    end

    # Loads constants associated with COM object _com_.  This is an internal
    # method that shouldn’t be called outside of this class.
    #
    # @param [WIN32OLE] com COM object to load constants from
    # @return [Boolean] Whether or not any constants where loaded
    def load_constants(com)
      return if constants_loaded?
      modul = nesting[-2]
      com.load_constants modul, typelib
      @constants_loaded = true
    end

    # Queries whether constants have already been loaded for this class.
    #
    # @return Whether or not constants have already been loaded for this class
    def constants_loaded?
      @constants_loaded ||= false
    end

  private

    # Gets the nesting of modules that lead up to this class.
    #
    # @return [Array<Module>] Modules that nest this class
    def nesting
      result = []
      name.split(/::/).inject(Module) do |modul, name|
        modul.const_get(name).tap{ |c| result << c }
      end
      result
    end
  end

  # Connects to or creates a new instance of a COM object.
  #
  # @option options [Boolean] :connect (false) Whether or not to connect to a
  #   running instance (see {.connect})
  # @option options [Boolean] :constants (true) Whether or not to load
  #   constants associated with the COM object (see {.constants})
  def initialize(options = {})
    @connected = false
    connect if options.fetch(:connect, self.class.connect?)
    self.com = COM.new(self.class.program_id) unless connected?
    self.class.load_constants(com) if
      options.fetch(:constants, self.class.constants?)
  end

  # Queries whether or not an already running COM object was connected to.
  # This is useful when deciding whether or not to close down the COM object
  # when you’re done with it.
  #
  # @return Whether or not an already running COM object was connected to
  def connected?
    @connected
  end

  def inspect
    '#<%s%s>' % [self.class, connected? ? ' connected' : '']
  end

private

  # Try to connect to a running COM object.
  def connect
    self.com = COM.connect(self.class.program_id)
    @connected = true
  rescue COM::OperationUnavailableError
  end
end
