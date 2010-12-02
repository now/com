# -*- coding: utf-8 -*-

module COM::PatternError
  # @private method used by {COM::Error.find}.
  def replaces?(error)
    pattern =~ error.message
  end

protected

  def replace(error)
    new(*pattern.match(error.message).captures)
  end

private

  def pattern(pattern = nil)
    (pattern or not defined? @pattern) ? @pattern = pattern : @pattern
  end
end
