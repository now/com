# -*- coding: utf-8 -*-

class COM::Events
  def initialize(object, interface)
    @observers = Hash.new{ [] }
    @events = WIN32OLE_EVENT.new(object, interface)
  end

  def register(*events)
    events.each do |event|
      @events.on_event event do |*args|
        @observers[event].each do |observer|
          observer.call(*args)
        end
      end
    end
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
