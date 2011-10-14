# -*- coding: utf-8 -*-

# Provides a simple wrapper around COM events.  WIN32OLE provides no way of
# releasing an observer and thus causes an unbreakable chain of objects,
# causing uncollectable garbage.  This class obviates most of the problems.
class COM::Events
  ArgumentMissing = Object.new.freeze

  # Creates a COM events wrapper for the COM object _com_’s _interface_.
  # Optionally, any _events_ may be given as additional arguments.
  #
  # @param [WIN32OLE] com COM object implementing _interface_
  # @param [String] interface Name of the COM interface that _com_ implements
  # @param [Array<String>] events Names of events to register (see
  #   {#register})
  def initialize(com, interface = nil)
    @interface = WIN32OLE_EVENT.new(com, interface)
    @events = {}
  end

  # Observe _event_ in _block_ during the execution of _during_.
  #
  # @param [String] event Name of the event to observe
  # @param [Proc] during Block during which to observe _event_
  # @yield [*args] Event arguments (specific for each event)
  # @return The result of _during_
  def observe(event, observer = ArgumentMissing)
    if ArgumentMissing.equal? observer
      register event, Proc.new
      return self
    end
    return self unless observer
    observer = observer.to_proc
    register event, observer
    return self unless block_given?
    begin
      yield
    ensure
      unobserve event, observer
    end
    self
  end

  def unobserve(event, observer = nil)
    @events[event].delete observer if observer
    if observer.nil? or @events[event].empty?
      @interface.off_event event
      @events.delete event
    end
    self
  end

  private

  def register(event, observer)
    if @events.include? event
      @events[event] << observer
      return
    end
    @interface.on_event event do |*args|
      @events[event].reduce({}){ |result, o|
        r = o.call(*args)
        result.merge! r if Hash === r
      }
    end
    @events[event] = [observer]
  end
end
