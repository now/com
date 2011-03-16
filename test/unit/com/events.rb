# -*- coding: utf-8 -*-

Expectations do
  expect WIN32OLE_EVENT.to.receive.new(:com, :interface) do
    COM::Events.new(:com, :interface)
  end

  expect mock.to.receive.on_event(:on_something) do |o|
    stub(WIN32OLE_EVENT).new{ o }
    COM::Events.new(:ole, :interface).register :on_something
  end

  expect mock.to.receive.call do |o|
    events = stub
    class << events
      def on_event(event, &block)
        @block = block
      end

      def trigger
        @block.call
      end
    end
    stub(WIN32OLE_EVENT).new{ events }
    e = COM::Events.new(:ole, :interface)
    e.register :on_something
    e.observe(:on_something, proc{ events.trigger }){ o.call }
  end
end
