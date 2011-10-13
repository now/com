# -*- coding: utf-8 -*-

class Pathname
  # Returns a String representation of this Pathname in the COM character set.
  #
  # @see COM.charset
  def to_com
    to_path.encode(COM.charset)
  end
end
