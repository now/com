# -*- coding: utf-8 -*-

require 'lookout'
require 'pathname'

require 'com'

Expectations do
  expect 'c:/a' do
    COM.stubs(:charset).returns('UTF-8')
    Pathname('c:/a').to_com
  end

  expect 'c:/å' do
    COM.stubs(:charset).returns('UTF-8')
    Pathname('c:/å').to_com
  end

  # ANSI and ISO-8859-1 å
  expect 'c:/å' do
    COM.stubs(:charset).returns('UTF-8')
    Pathname("c:/\xe5").to_com
  end
end
