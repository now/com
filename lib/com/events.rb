# -*- coding: utf-8 -*-

# Provides a simple wrapper around COM events.  WIN32OLE provides no way of
# releasing an observer and thus causes an unbreakable chain of objects,
# causing uncollectable garbage.  This class obviates most of the problems.
class COM::Events
  # Creates a COM events wrapper for the COM object _com_’s _interface_.
  # Optionally, any _events_ may be given as additional arguments.
  #
  # @param [WIN32OLE] com COM object implementing _interface_
  # @param [String] interface Name of the COM interface that _com_ implements
  # @param [Array<String>] events Names of events to register (see
  #   {#register})
  def initialize(com, interface, *events)
    @observers = Hash.new{ [] }
    @events = WIN32OLE_EVENT.new(com, interface)
    register(*events)
  end

  # Register _events_.
  #
  # @param [Array<String>] events Names of events to register
  # @return self
  def register(*events)
    events.each do |event|
      saved_verbose, $VERBOSE = $VERBOSE, nil
      begin
        @events.on_event event do |*args|
          @observers[event].each do |observer|
            observer.call(*args)
          end
        end
      ensure
        $VERBOSE = saved_verbose
      end
    end
    self
  end

  # Observe _event_ in _block_ during the execution of _during_.
  #
  # @param [String] event Name of the event to observe
  # @param [Proc] during Block during which to observe _event_
  # @yield [*args] Event arguments (specific for each event)
  # @return The result of _during_
  def observe(event, during, &block)
    @observers[event] <<= block
    begin
      during.call
    ensure
      @observers[event].delete block
    end
  end
end
