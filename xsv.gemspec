
lib = File.expand_path("../lib", __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require "xsv/version"

Gem::Specification.new do |spec|
  spec.name          = "xsv"
  spec.version       = Xsv::VERSION
  spec.authors       = ["Martijn Storck"]
  spec.email         = ["martijn@storck.io"]

  spec.summary       = "Minimal xlsx parser that provides nothing a CSV parser wouldn't"
  spec.homepage      = "https://github.com/martijn/xsv"
  spec.license       = "MIT"

  if spec.respond_to?(:metadata)
    spec.metadata["homepage_uri"] = spec.homepage
    spec.metadata["source_code_uri"] = "https://github.com/martijn/xsv"
    spec.metadata["changelog_uri"] = "https://github.com/martijn/xsv/CHANGELOG.md"
  else
    raise "RubyGems 2.0 or newer is required to protect against " \
      "public gem pushes."
  end

  # Specify which files should be added to the gem when it is released.
  # The `git ls-files -z` loads the files in the RubyGem that have been added into git.
  spec.files         = Dir.chdir(File.expand_path('..', __FILE__)) do
    `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  end
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.required_ruby_version = '~> 2.6'

  spec.add_dependency "rubyzip", "~> 2.2.0"
  spec.add_dependency "nokogiri", "~> 1.10.8"

  spec.add_development_dependency "bundler", "~> 1.17"
  spec.add_development_dependency "rake", "~> 10.0"
  spec.add_development_dependency "minitest", "~> 5.0"
end
