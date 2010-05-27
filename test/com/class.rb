# -*- coding: utf-8 -*-

require 'lookout'

require 'com'

Expectations do
  expect 'Program.Class' do
    Class.new(COM::Class).program_id('Program.Class')
  end

  expect 'Program.Class' do
    Class.new(COM::Class){
      program_id 'Program.Class'
    }.program_id
  end

  expect ArgumentError do
    Class.new(COM::Class){
      stubs(:name).returns('Class')
    }.program_id
  end

  expect 'Program.Class' do
    Class.new(COM::Class){
      stubs(:name).returns('Program::Class')
    }.program_id
  end

  expect 'Program.Class' do
    Class.new(COM::Class){
      stubs(:name).returns('Vendor::Program::Class')
    }.program_id
  end

  expect false do
    Class.new(COM::Class).connect?
  end

  expect true do
    Class.new(COM::Class).connect
  end

  expect true do
    Class.new(COM::Class){
      connect
    }.connect?
  end

  expect true do
    Class.new(COM::Class).constants?
  end

  expect true do
    Class.new(COM::Class).constants(true)
  end

  expect true do
    Class.new(COM::Class){
      constants true
    }.constants?
  end

  expect COM.to.receive(:connect).with('A.B').returns(ignore) do
    Class.new(COM::Class){ program_id 'A.B' }.new(:connect => true, :constants => false)
  end

  expect COM.to.receive(:new).never do
    COM.stubs(:connect).returns(ignore)
    Class.new(COM::Class){ program_id 'A.B' }.new(:connect => true, :constants => false)
  end

  expect COM.to.receive(:new).with('A.B').returns(ignore) do
    COM.stubs(:connect).raises(COM::OperationUnavailableError)
    Class.new(COM::Class){ program_id 'A.B' }.new(:connect => true, :constants => false)
  end

  expect COM.to.receive(:new).with('A.B').returns(ignore) do
    COM.stubs(:connect).raises(COM::OperationUnavailableError)
    Class.new(COM::Class){ program_id 'A.B' }.new(:constants => false)
  end

  expect true do
    COM.stubs(:connect).returns(ignore)
    Class.new(COM::Class){ program_id 'A.B' }.new(:connect => true, :constants => false).connected?
  end

  expect false do
    COM.stubs(:connect).raises(COM::OperationUnavailableError)
    COM.stubs(:new).returns(ignore)
    Class.new(COM::Class){ program_id 'A.B' }.new(:connect => true, :constants => false).connected?
  end

  expect false do
    COM.stubs(:new).returns(ignore)
    Class.new(COM::Class){ program_id 'A.B' }.new(:constants => false).connected?
  end

  expect :result do
    COM.stubs(:new).returns(ignore)
    Class.new(COM::Class){ program_id 'A.B' }.new(:constants => false).with_properties({}){ :result }
  end

  expect 1 do
    o = {:a => 1}
    COM.stubs(:new).returns(o)
    Class.new(COM::Class){ program_id 'A.B' }.new(:constants => false).with_properties(:a => 2){ }
    o[:a]
  end

  expect 1 do
    o = {:a => 1}
    COM.stubs(:new).returns(o)
    Class.new(COM::Class){ program_id 'A.B' }.new(:constants => false).with_properties(:a => 2){ raise } rescue nil
    o[:a]
  end

  expect 2 do
    o = {:a => 1}
    COM.stubs(:new).returns(o)
    v = nil
    Class.new(COM::Class){ program_id 'A.B' }.new(:constants => false).with_properties(:a => 2){ v = o[:a] }
    v
  end

  expect StandardError do
    COM.stubs(:new).returns(ignore)
    Class.new(COM::Class){ program_id 'A.B' }.new(:constants => false).with_properties({}){ raise StandardError }
  end

  expect StandardError do
    o = stub_everything
    o.stubs(:[]=).returns(:whatever).then.raises(StandardError)
    COM.stubs(:new).returns(o)
    Class.new(COM::Class){ program_id 'A.B' }.new(:constants => false).with_properties(:a => 2){ raise StandardError }
  end

  expect(Class.new(COM::Class){ program_id 'A.B' }.to.receive(:load_constants)) do |o|
    COM.stubs(:new).returns(ignore)
    o.new
  end

  expect(Class.new(COM::Class){ program_id 'A.B' }.to.receive(:load_constants).never) do |o|
    COM.stubs(:new).returns(ignore)
    o.new(:constants => false)
  end
end
