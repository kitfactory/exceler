# Exceler

Excel document parser for project metrics.
プロジェクト計測に使用できる、エクセル文書パーサー

## Installation

    $ gem install exceler

## Usage

```
require 'exceler'

#find Excel files エクセルファイルを探します。
files = list = Exceler.list_files( "." )

#create item scan option　エクセルファイルをスキャンするオプションを設定します。
# sheet,header,id,contnet,assign,start,limit,state,state_condition
so = Exceler::ScanOption.new( nil, 1 ,  "B" , "C",  "D" ,  nil , nil , "E" , nil )

#find items in excel files エクセルファイルからアイテムを探します。
items = Exceler.scan_items( list ,so )
	
#find persons who are assigned to some issues.　アイテムの担当をリストアップします。
plist = Exceler.list_assigned_person( items )
	
#create issue list for each person　担当ごとに作業します。
s = "" # HTML
m = {};	

for p in plist
	pi = Exceler.pickup_assigned( items , p ) #担当に割り当てられたアイテムを取得します。
	pi = Exceler.pickup_incomplete( pi ) # そのうち、未完了のものを取得します。
	if( pi.length != 0 ) # 未完了なものがあればHTMLにエクスポートしておきます。
		s+=Exceler.export_item_html( pi , p ,nil )
		m[p]=pi.length.to_s	# 担当ごとの残アイテム数をMAPにしておきます。
	end
end
	
#export as html	
Exceler.write_html_file( "test.html" , s , nil )

#export as csv
Exceler.write_csv_file( "test.csv" , m )

```

## Contributing

1. Fork it ( https://github.com/kitfactory/exceler/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request
