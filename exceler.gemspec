# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'exceler/version'

Gem::Specification.new do |spec|
  spec.name          = "exceler"
  spec.version       = Exceler::VERSION
  spec.authors       = ["kitfactory"]
  spec.email         = ["kitfactory@gmail.com"]
  spec.summary       = %q{Excel document parser for project metrics.}
  spec.description   = %q{check homepage.}
  spec.homepage      = "https://github.com/kitfactory/exceler"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.6"
  spec.add_development_dependency "rake"
  spec.add_dependency "roo"
end
