# -*- coding: utf-8 -*-

require 'lookout'

require 'com'

Expectations do
  expect WIN32OLE_EVENT.to.receive(:new).with(:object, :interface) do
    COM::Events.new(:object, :interface)
  end

  expect mock.to.receive(:on_event).with(:on_something) do |o|
    WIN32OLE_EVENT.stubs(:new).returns(o)
    COM::Events.new(:ole, :interface).register :on_something
  end

  expect mock.to.receive(:call) do |o|
    events = stub_everything
    class << events
      def on_event(event, &block)
        @block = block
      end

      def trigger
        @block.call
      end
    end
    WIN32OLE_EVENT.stubs(:new).returns(events)
    e = COM::Events.new(:ole, :interface)
    e.register :on_something
    e.observe(:on_something, proc{ events.trigger }){ o.call }
  end
end
