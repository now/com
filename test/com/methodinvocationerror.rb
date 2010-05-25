# -*- coding: utf-8 -*-

require 'lookout'

require 'com'

Expectations do
  M = stub :message => <<EOM
method
  OLE error code:123abcDE in server
    message
  HRESULT error code:0x123abcDE
    hresult message
EOM

  expect COM::MethodInvocationError do
    COM::Error.from(M)
  end

  expect 'method' do
    COM::Error.from(M).method
  end

  expect 'server' do
    COM::Error.from(M).server
  end

  expect 0x123abcDE do
    COM::Error.from(M).error_code
  end

  expect 0x123abcDE do
    COM::Error.from(M).hresult_error_code
  end

  expect 'hresult message' do
    COM::Error.from(M).hresult_message
  end

  expect 'message' do
    COM::Error.from(M).message
  end

  expect 'server:method:message (0x123abcde)' do
    COM::Error.from(M).to_s
  end
end
