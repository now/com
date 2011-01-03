# -*- coding: utf-8 -*-

class Pathname
  # Returns a String representation of this Pathname in the COM character set.
  #
  # @see COM.charset
  def to_com
    require 'iconv' unless Object.const_defined? :Iconv
    ['UTF-8', 'MS-ANSI', 'ISO-8859-1'].map{ |from|
      Iconv.conv(COM.charset, from, to_str) rescue nil
    }.find{ |path| path } || to_str
  end
end
