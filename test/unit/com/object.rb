# -*- coding: utf-8 -*-

Expectations do
=begin
  expect :result do
    COM::Object.new(stub).with_properties({}){ :result }
  end

  expect 1 do
    o = {:a => 1}
    COM::Object.new(o).with_properties(:a => 2){ }
    o[:a]
  end

  expect 1 do
    o = {:a => 1}
    COM::Object.new(o).with_properties(:a => 2){ raise } rescue nil
    o[:a]
  end

  expect 2 do
    o = {:a => 1}
    v = nil
    COM::Object.new(o).with_properties(:a => 2){ v = o[:a] } rescue nil
    o[:a]
    v
  end
=end

  expect StandardError do
    COM::Object.new(stub).with_properties({}){ raise StandardError }
  end
end
