# -*- coding: utf-8 -*-

$:.unshift File.expand_path("../lib", __FILE__)

require 'rake/gempackagetask'
require 'rake/testtask'
require 'rubygems/dependency_installer'
require 'yard'
require 'com/version'

task :default => :test

specification = Gem::Specification.new do |s|
  s.name   = 'com'
  s.summary = 'An object-oriented wrapper around WIN32OLE'
  s.version = COM::Version
  s.author = 'Nikolai Weibull'
  s.email = 'now@bitwi.se'
  s.homepage = 'http://github.com/now/com'
  s.description = <<EOD
COM is an object-oriented wrapper around WIN32OLE that makes it easier
to add behavior to WIN32OLE objects.
EOD

  s.files = FileList['{lib,test}/**/*.rb', '[A-Z]*']

  s.add_development_dependency 'yard', '>= 0.2.3.5'
end

Rake::TestTask.new do |t|
  t.libs << 'test'
  t.pattern = 'test/**/*.rb'
end

YARD::Rake::YardocTask.new

Rake::GemPackageTask.new(specification) do |g|
  desc 'Run :package and install the resulting gem'
  task :install => :package do
    Gem::DependencyInstaller.new.install File.join(g.package_dir, g.gem_file)
  end
end
