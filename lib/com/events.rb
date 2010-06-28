# -*- coding: utf-8 -*-

class COM::Events
  def initialize(com, interface, *events)
    @observers = Hash.new{ [] }
    @events = WIN32OLE_EVENT.new(com, interface)
    register(*events)
  end

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

  def observe(event, during, &block)
    @observers[event] <<= block
    begin
      during.call
    ensure
      @observers[event].delete block
    end
  end
end
