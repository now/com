# -*- coding: utf-8 -*-

require 'pathname'

Expectations do
  expect 'c:/a' do
    stub(COM).charset{ 'UTF-8' }
    Pathname('c:/a').to_com
  end

  expect 'c:/å' do
    stub(COM).charset{ 'UTF-8' }
    Pathname('c:/å').to_com
  end

  # ANSI and ISO-8859-1 å
  expect 'c:/å' do
    stub(COM).charset{ 'UTF-8' }
    Pathname("c:/\xe5".force_encoding('Windows-1252')).to_com
  end
end
