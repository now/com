# -*- coding: utf-8 -*-

class COM::Error < RuntimeError
  def self.from(error)
    errors.find{ |replacement| replacement.replaces? error }.replace(error)
  end

  def self.replaces?(error)
    true
  end

  def self.replace(error)
    new error.message
  end

  def self.inherited(error)
    errors.unshift error
  end
  private_class_method :inherited

  def self.errors
    @@errors ||= [self]
  end
  private_class_method :errors
end
