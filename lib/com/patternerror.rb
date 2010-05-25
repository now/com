# -*- coding: utf-8 -*-

module COM::PatternError
  def pattern(pattern = nil)
    return @pattern if instance_variable_defined? :@pattern and pattern.nil?
    @pattern = pattern
  end
  private :pattern

  def replaces?(error)
    pattern =~ error.message
  end

  def replace(error)
    new(*pattern.match(error.message).captures)
  end
end
