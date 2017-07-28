# frozen_string_literal: true

lib = File.expand_path('../lib/', __FILE__)
$LOAD_PATH.unshift lib unless $LOAD_PATH.include?(lib)

Gem::Specification.new do |s|
  s.name        = 'ruby_xl_poi'

  path = File.expand_path('lib/ruby_xl_poi/version.rb', File.dirname(__FILE__))
  s.version     = File.read(path).match(/\s*VERSION\s*=\s*['"](.*)['"]/)[1]
  s.date        = '2017-07-27'
  s.summary     = 'Modify excel sheets using the powerfull Apache POI'
  s.description = 'There are quite a few pure ruby gems to deal with excel files
                   with various degree of success. This gem\'s approach is to
                   use powerfull software like Apache POI and simply provide
                   ruby like interface.'
  s.files       = 'ruby_xl_poi.rb'
  s.authors     = ['Andrea Salicetti']
  s.email       = 'andrea.salicetti@gmail.com'
  s.homepage    = 'http://github.com/knightq/ruby_xl_poi'
  s.license     = 'MIT'

  s.required_ruby_version = '~> 2.0'
  s.add_runtime_dependency 'rjb', ['~> 1.4']

  s.add_development_dependency 'rake'

  s.require_path = 'lib'
  s.files        = Dir.glob('{lib}/**/*') + [
    'ruby_xl_poi.gemspec',
    'README.md'
  ]
end
