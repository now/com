# -*- coding: utf-8 -*-

require 'lookout'

require 'com'

Expectations do
  expect 'UTF-8' do
    WIN32OLE.stubs(:codepage).returns(WIN32OLE::CP_UTF8)
    COM.charset
  end

  expect RuntimeError do
    WIN32OLE.stubs(:codepage).returns(:unknown_encoding)
    COM.charset
  end

  expect COM::Error do
    WIN32OLE.stubs(:connect).raises(WIN32OLERuntimeError)
    COM.connect(ignore)
  end

  expect COM::Error do
    WIN32OLE.stubs(:new).raises(WIN32OLERuntimeError)
    COM.new(ignore)
  end
end
