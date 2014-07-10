# Exceler

Excel document parser for project metrics.

## Installation

    $ gem install exceler

## Usage

```
require 'exceler'

# Example1 F列が埋まっていれば済とみなす例

list = Exceler.list_files( "test1" )
so = ScanOption.new( 0 , "A" , "B" , "D" , "E" , "F" , nil )
items = Exceler.scan_items( list ,so )

# Example2 F列が済となっていれば済とみなす例

list = Exceler.list_files( "test2" )
so = ScanOption.new( 2 , "A" , "B" , "D" , "E" , "F" , "済" )
items = Exceler.scan_items( list ,so )
assigned_items = Exceler.pickup_assigned( items , "山田")
incomplete_items = Exceler.pickup_incomplete( items )
expiration_items = Exceler.pickup_expiration( items )

puts "Total items:" + items.length.to_s
puts "Yamda assgined items :" + assigned_items.length.to_s
puts "Incomplete items :" + incomplete_items.length.to_s
puts "Expiration items :" + expiration_items.length.to_s

```

## Contributing

1. Fork it ( https://github.com/kitfactory/exceler/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request
