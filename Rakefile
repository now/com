# -*- coding: utf-8 -*-

require 'lookout/rake/tasks'
require 'yard'

Lookout::Rake::Tasks::Test.new
Lookout::Rake::Tasks::Gem.new
YARD::Rake::YardocTask.new do |t|
  t.options.concat %w[--markup markdown]
end
