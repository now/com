# -*- coding: utf-8 -*-

Expectations do
  expect COM::Error do
    COM::Error.from(stub(:message => '<some random message>'))
  end

  expect '<some random message>' do
    COM::Error.from(stub(:message => '<some random message>')).message
  end
end
