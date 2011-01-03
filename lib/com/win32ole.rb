# -*- coding: utf-8 -*-

class WIN32OLE_TYPE
  class << self
    def enums(id)
      ole_classes(id).select{ |c| c.visible? and c.ole_type == 'Enum' }
    end
  end

  def constants
    variables.select{ |v| v.visible? and v.variable_kind == 'CONSTANT' }
  end
end

class WIN32OLE_VARIABLE
  def const_name
    name.sub(/^./){ |l| l.upcase }
  end
end
