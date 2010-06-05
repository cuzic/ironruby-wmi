dir = File.dirname(__FILE__)
unless $:.include?(dir) || $:.include?(File.expand_path(dir))
  $:.unshift(dir)
end
require 'lib/ironruby-wmi.rb'
