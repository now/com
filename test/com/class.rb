# -*- coding: utf-8 -*-

require 'lookout'

require 'com'

Expectations do
  expect :result do
    COM::Class.new(ignore).with_properties({}){ :result }
  end

  expect 1 do
    o = {:a => 1}
    COM::Class.new(o).with_properties(:a => 2){ }
    o[:a]
  end

  expect 1 do
    o = {:a => 1}
    COM::Class.new(o).with_properties(:a => 2){ raise } rescue nil
    o[:a]
  end

  expect 2 do
    o = {:a => 1}
    v = nil
    COM::Class.new(o).with_properties(:a => 2){ v = o[:a] } rescue nil
    o[:a]
    v
  end

  expect StandardError do
    COM::Class.new(ignore).with_properties({}){ raise StandardError }
  end

  expect StandardError do
    o = stub_everything
    o.stubs(:[]=).returns(:whatever).then.raises(StandardError)
    COM::Class.new(o).with_properties(:a => 2){ }
  end
end
