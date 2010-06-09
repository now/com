# -*- coding: utf-8 -*-

module COM::MethodMissing
  def method_missing(*args)
    super
  rescue WIN32OLERuntimeError => e
    raise COM::Error.from(e)
  end
  private :method_missing
end
