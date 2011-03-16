# -*- coding: utf-8 -*-

Expectations do
  expect 'Program.Class' do
    Class.new(COM::Instantiable).program_id('Program.Class')
  end

  expect 'Program.Class' do
    Class.new(COM::Instantiable){
      program_id 'Program.Class'
    }.program_id
  end

  expect ArgumentError do
    Class.new(COM::Instantiable).tap{ |c| stub(c).name{ 'Class' } }.program_id
  end

  expect 'Program.Class' do
    Class.new(COM::Instantiable).tap{ |c| stub(c).name{ 'Program::Class' } }.program_id
  end

  expect 'Program.Class' do
    Class.new(COM::Instantiable).tap{ |c| stub(c).name{ 'Vendor::Program::Class' } }.program_id
  end

  expect false do
    Class.new(COM::Instantiable).connect?
  end

  expect true do
    Class.new(COM::Instantiable).connect
  end

  expect true do
    Class.new(COM::Instantiable){ connect }.connect?
  end

  expect true do
    Class.new(COM::Instantiable).constants?
  end

  expect true do
    Class.new(COM::Instantiable).constants(true)
  end

  expect true do
    Class.new(COM::Instantiable){ constants true }.constants?
  end

  expect COM.to.receive.connect('A.B'){ stub } do
    Class.new(COM::Instantiable){ program_id 'A.B' }.new(:connect => true, :constants => false)
  end

  expect COM.not.to.receive.new do
    stub(COM).connect{ stub }
    Class.new(COM::Instantiable){ program_id 'A.B' }.new(:connect => true, :constants => false)
  end

  expect COM.to.receive.new('A.B'){ stub } do
    stub(COM).connect{ raise COM::OperationUnavailableError }
    Class.new(COM::Instantiable){ program_id 'A.B' }.new(:connect => true, :constants => false)
  end

  expect COM.to.receive.new('A.B'){ stub } do
    stub(COM).connect{ raise COM::OperationUnavailableError }
    Class.new(COM::Instantiable){ program_id 'A.B' }.new(:constants => false)
  end

  expect true do
    stub(COM).connect{ stub }
    Class.new(COM::Instantiable){ program_id 'A.B' }.new(:connect => true, :constants => false).connected?
  end

  expect false do
    stub(COM).connect{ raise COM::OperationUnavailableError }
    stub(COM).new{ stub }
    Class.new(COM::Instantiable){ program_id 'A.B' }.new(:connect => true, :constants => false).connected?
  end

  expect false do
    stub(COM).new{ stub }
    Class.new(COM::Instantiable){ program_id 'A.B' }.new(:constants => false).connected?
  end

  expect(Class.new(COM::Instantiable){ program_id 'A.B' }.to.receive.load_constants) do |o|
    stub(COM).new{ stub }
    o.new
  end

  expect(Class.new(COM::Instantiable){ program_id 'A.B' }.not.to.receive.load_constants) do |o|
    stub(COM).new{ stub }
    o.new(:constants => false)
  end
end
