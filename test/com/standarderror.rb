# -*- coding: utf-8 -*-

Expectations do
  expect NotImplementedError do
    COM::Error.from(stub(:message => "\nHRESULT error code:0x80004001\n"))
  end

  expect NoMethodError do
    COM::Error.from(stub(:message => "\nHRESULT error code:0x80020006\n"))
  end

  expect ArgumentError do
    COM::Error.from(stub(:message => "\nHRESULT error code:0x8002000e\n"))
  end

  expect ArgumentError do
    COM::Error.from(stub(:message => "\nHRESULT error code:0x800401e4\n"))
  end

  expect COM::OperationUnavailableError do
    COM::Error.from(stub(:message => "\nHRESULT error code:0x800401e3\n"))
  end
end
