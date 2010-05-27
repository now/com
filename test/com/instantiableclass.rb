# -*- coding: utf-8 -*-

require 'lookout'

require 'com'

Expectations do
  expect 'Program.Class' do
    Class.new(COM::InstantiableClass).program_id('Program.Class')
  end

  expect 'Program.Class' do
    Class.new(COM::InstantiableClass){
      program_id 'Program.Class'
    }.program_id
  end

  expect ArgumentError do
    Class.new(COM::InstantiableClass){
      stubs(:name).returns('Class')
    }.program_id
  end

  expect 'Program.Class' do
    Class.new(COM::InstantiableClass){
      stubs(:name).returns('Program::Class')
    }.program_id
  end

  expect 'Program.Class' do
    Class.new(COM::InstantiableClass){
      stubs(:name).returns('Vendor::Program::Class')
    }.program_id
  end

  expect false do
    Class.new(COM::InstantiableClass).connect?
  end

  expect true do
    Class.new(COM::InstantiableClass).connect
  end

  expect true do
    Class.new(COM::InstantiableClass){
      connect
    }.connect?
  end

  expect true do
    Class.new(COM::InstantiableClass).constants?
  end

  expect true do
    Class.new(COM::InstantiableClass).constants(true)
  end

  expect true do
    Class.new(COM::InstantiableClass){
      constants true
    }.constants?
  end

  expect COM.to.receive(:connect).with('A.B').returns(ignore) do
    Class.new(COM::InstantiableClass){ program_id 'A.B' }.new(:connect => true, :constants => false)
  end

  expect COM.to.receive(:new).never do
    COM.stubs(:connect).returns(ignore)
    Class.new(COM::InstantiableClass){ program_id 'A.B' }.new(:connect => true, :constants => false)
  end

  expect COM.to.receive(:new).with('A.B').returns(ignore) do
    COM.stubs(:connect).raises(COM::OperationUnavailableError)
    Class.new(COM::InstantiableClass){ program_id 'A.B' }.new(:connect => true, :constants => false)
  end

  expect COM.to.receive(:new).with('A.B').returns(ignore) do
    COM.stubs(:connect).raises(COM::OperationUnavailableError)
    Class.new(COM::InstantiableClass){ program_id 'A.B' }.new(:constants => false)
  end

  expect true do
    COM.stubs(:connect).returns(ignore)
    Class.new(COM::InstantiableClass){ program_id 'A.B' }.new(:connect => true, :constants => false).connected?
  end

  expect false do
    COM.stubs(:connect).raises(COM::OperationUnavailableError)
    COM.stubs(:new).returns(ignore)
    Class.new(COM::InstantiableClass){ program_id 'A.B' }.new(:connect => true, :constants => false).connected?
  end

  expect false do
    COM.stubs(:new).returns(ignore)
    Class.new(COM::InstantiableClass){ program_id 'A.B' }.new(:constants => false).connected?
  end

  expect(Class.new(COM::InstantiableClass){ program_id 'A.B' }.to.receive(:load_constants)) do |o|
    COM.stubs(:new).returns(ignore)
    o.new
  end

  expect(Class.new(COM::InstantiableClass){ program_id 'A.B' }.to.receive(:load_constants).never) do |o|
    COM.stubs(:new).returns(ignore)
    o.new(:constants => false)
  end
end
