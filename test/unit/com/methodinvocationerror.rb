# -*- coding: utf-8 -*-

Expectations do
  expect COM::MethodInvocationError do
    COM::Error.from(stub(:message => %{method
  OLE error code:123abcDE in server
    message
  HRESULT error code:0x123abcDE
    hresult message}))
  end

  expect 'method' do
    COM::Error.from(stub(:message => %{method
  OLE error code:123abcDE in server
    message
  HRESULT error code:0x123abcDE
    hresult message})).method
  end

  expect 'server' do
    COM::Error.from(stub(:message => %{method
  OLE error code:123abcDE in server
    message
  HRESULT error code:0x123abcDE
    hresult message})).server
  end

  expect 0x123abcDE do
    COM::Error.from(stub(:message => %{method
  OLE error code:123abcDE in server
    message
  HRESULT error code:0x123abcDE
    hresult message})).code
  end

  expect 0x123abcDE do
    COM::Error.from(stub(:message => %{method
  OLE error code:123abcDE in server
    message
  HRESULT error code:0x123abcDE
    hresult message})).hresult_code
  end

  expect 'hresult message' do
    COM::Error.from(stub(:message => %{method
  OLE error code:123abcDE in server
    message
  HRESULT error code:0x123abcDE
    hresult message})).hresult_message
  end

  expect 'message' do
    COM::Error.from(stub(:message => %{method
  OLE error code:123abcDE in server
    message
  HRESULT error code:0x123abcDE
    hresult message})).message
  end

  expect 'message' do
    COM::Error.from(stub(:message => %{method
  OLE error code:123abcDE in server
    message
  HRESULT error code:0x123abcDE
    hresult message})).to_s
  end
end
