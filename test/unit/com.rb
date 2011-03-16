# -*- coding: utf-8 -*-

Expectations do
  expect 'UTF-8' do
    stub(WIN32OLE).codepage{ WIN32OLE::CP_UTF8 }
    COM.charset
  end

  expect RuntimeError do
    stub(WIN32OLE).codepage{ :unknown_encoding }
    COM.charset
  end

  expect COM::Error do
    stub(WIN32OLE).connect{ raise WIN32OLERuntimeError }
    COM.connect(stub)
  end

  expect COM::Error do
    stub(WIN32OLE).new{ raise WIN32OLERuntimeError }
    COM.new(stub)
  end
end
