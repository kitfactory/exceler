require "exceler/version"
require "roo"

module Exceler
#
# ScanOption
# アイテムを取得する際のオプション
# 
class ScanOption

	#
	# new
	# 
	# ==== Args
	# sheet :: シート名(nilの場合は全てのシートに適用)
	# header :: ヘッダー行(スキップする行数)
	# id_row :: アイテムの存在を確認する列
	# content_row :: コンテンツの内容を表す列
	# assign_row :: 担当者の列
	# start_row :: 開始日の列
	# limit_row :: 期限の列
	# state_row :: ステータスの列
	# state_condition :: 合致で済とする場合は合致の文字列、埋まっていることで済とする場合はnil
	# 
	def initialize( sheet, header , id_row , content_row , assign_row , start_row , limit_row , state_row,  state_condition )
		@sheet = sheet
		@header = header			# ヘッダー行(スキップする行数)
		@id_row = id_row			# アイテムの存在を確認する列
		@content_row = content_row
		@assign_row	= assign_row	# 担当者の列
		@start_row	= start_row		# 開始日の列
		@limit_row	= limit_row		# 期限の列
		@state_row	= state_row		# ステータスの列
		@state_condition = state_condition
	end

	attr_reader :sheet
	attr_reader :header
	attr_reader :id_row
	attr_reader	:content_row
	attr_reader :assign_row
	attr_reader :start_row
	attr_reader :limit_row
	attr_reader :state_row
	attr_reader :state_condition
end

#
# 実施状況を確認するアイテム
#
class Item
	COMPLETE = 1
	INCOMPLETE = 2

	def initialize
		@file = nil
		@id = nil
		@content = nil
		@state = INCOMPLETE
		@assign = nil
		@start = nil
		@limit = nil
	end

	attr_accessor :file
	attr_accessor :id
	attr_accessor :content
	attr_accessor :state
	attr_accessor :assign
	attr_accessor :start
	attr_accessor :limit
end

	XLS = "xls"
	XLSX = "xlsx"
	EXT_PATTERNS = [ XLS , XLSX ];

	# 指定されたディレクトリからファイル(.xls,.xlsx)を取得します
	# find Excel files from the specified directory
	# ==== Args
	# dir :: エクセルファイルを検索するディレクトリ
	# ==== Return
	# エクセルファイルの名前の配列
	def self.list_files( dir )
		ret = [];
		for ext in EXT_PATTERNS
			filepattern = dir+File::SEPARATOR+"*."+ext;
			Dir[filepattern].each do |file|  
#				puts "founds " + file
				ret.push( file )
			end
		end
		return ret
	end

 	#
 	#
 	#
 	def self.show_item( item )
 		s = ""
 		if( item.assign != nil )
 			s += ("assign:" + item.assign + "," )
 		end
 		if( item.start != nil )
 			s += ( "start:" + item.start.strftime("%Y/%m/%d") + "," )
 		end
 		if( item.limit != nil )
 			s += ( "limit:" + item.limit.strftime("%Y/%m/%d") + "," )
 		end
 		if( item.state != nil )
 			if( item.state == Item::COMPLETE )
 				s += ( "state: complete ")
 			else
 				s += "state: incomplete"
 			end
 		end
 		puts s
 	end

 	#
 	# オプションにしたがってファイルを解析し、アイテムを返却します。
 	# scan the items with specified option from the files
 	#
	# ==== Args
	# files :: エクセルファイルの配列
	# opt :: 検索時のオプション、ScanOptionオブジェクト
	# ==== Return
	# アイテムの配列
	def self.scan_items( files , opt )
		ret = []
		if( opt == nil )
			return nil
		end
		for file in files
			puts file
			re = Regexp.new( XLS+"$" )
			if( file =~ re ) # XLS file
				# puts "XLS file scan " + file
				s = Roo::Excel.new(file)
			else	#XLSX file
				# puts "XLSX file scan"+file
				s = Roo::Excelx.new(file) 
			end

			for sheet in s.sheets
				# if sheet option is nil then scan all sheets
				# else scan only one sheet that has the specified sheet name. 
				if( opt.sheet != nil )
					if( opt.sheet != sheet )
						next
					end
				end

				s.default_sheet = sheet
				if( s.first_row == nil )
					next
				else
					header = s.first_row
				end
				if( opt.header >= header )
					header = opt.header
				end
				(header..s.last_row).each do |num|
					c = s.cell( opt.id_row , num )
					if( c != nil )
						i = Item.new
						i.file = file
						i.id = s.cell( opt.id_row, num )
						if( opt.content_row != nil )
							i.content = s.cell( opt.content_row ,num )
						end
						if( opt.assign_row != nil )
							i.assign = s.cell( opt.assign_row , num )
							i.assign.strip!
						end
						if( opt.start_row != nil )
							i.start = s.cell( opt.start_row , num )
						end
						if( opt.limit_row != nil )
							i.limit = s.cell( opt.limit_row , num )
						end
						if( opt.state_row != nil )
							puts opt.state_row
							if( opt.state_condition == nil )
								if( s.cell( opt.state_row , num ) != nil )
									i.state = Item::COMPLETE
								else
									i.state = Item::INCOMPLETE
								end
							else
								if( s.cell( opt.state_row ,num ) == opt.state_condition )
									i.state = Item::COMPLETE
								else
									i.state = Item::INCOMPLETE
								end
							end
						end
						# show_item( i )
						ret.push( i )
					end
				end
			end
		end
		return ret
	end


	#
	#  渡されたアイテムのうち、割り当てられた人を一覧します。
	#  list item assigned person.
	#
	# ==== Args
	# items :: アイテムの配列
	# ==== Return
	# 担当に割あたっている人の配列
	#        
	def self.list_assigned_person( items )
		pl = {}
		for item in items
			if( item.assign != nil )
				pl[item.assign] = item
			end
		end
		return pl.keys
	end
	
	#
	#  渡されたアイテムのうち、特定の人に割り当てられたアイテムをピックアップします。
	#  pickup specified person assigned items from the specified items
	#
	# ==== Args
	# items :: アイテムの配列
	# assign :: 担当
	# ==== Return
	# 担当に割あたっているアイテムの配列
	#        
	def self.pickup_assigned( items , assign )
		ret = []
		for item in items
			if( item.assign == assign )
				ret.push( item )
			end
		end
		return ret
	end

	#
	# 渡されたアイテムのうち未完了のアイテムをピックアップします。
	# pickup incompleted items from the specified items
	#
	# ==== Args
	# items :: アイテムの配列
	# ==== Return
	# 未完了アイテムの配列
	def self.pickup_incomplete( items )
		ret = []
		for item in items
			if( item.state == Item::INCOMPLETE )
				ret.push( item )
			end
		end
		return ret
	end

	#
	# 期限切れのアイテムを探します
	# pickup limit exceeded items from the specified items
	#
	# ==== Args
	# items :: アイテムの配列
	# ==== Return
	# 期限切れになっているアイテムの配列
	def self.pickup_expired( items )
		ret = []
		current = Date.today
		incomplete = pickup_incomplete( items )
		for item in incomplete
			if( item.limit != nil )
				# puts item.limit.strftime("%Y/%m/%d")+"-"+current.strftime("%Y/%m%d")
				if( item.limit < current )
					ret.push(item)
				end
			end
		end
		return ret
	end

	#
	# タスクの状況をHTMLにする。
	# transform task status to html string.
	#
	# ==== Args
	# items :: アイテムの配列
	# ==== Return
	# HTML Content
	#
	def self.export_item_html( items , title , subtitle )
		s=""
		if( title != nil )
			s+="<p class='exceler-title'><h3>"+title
			if( subtitle != nil )
				s+=+"-" + subtitle+ ":" + items.length.to_s
			end
			s+="</h3><br>"
		end
		s += "<table class='exceler-table'>"
		for item in items
			s+="<tr>"
			if( item.file != nil )
				s+=( "<td>"+ item.file + "</td>" )
			end
			if( item.content != nil )
				s+=( "<td>"+ item.content.to_s + "</td>" )
			end
			if( item.limit != nil )
				s+=( "<td>"+ item.limit.to_s + "</td>" )
			end
			s+="</td>"
		end
		s+= "</table>"
		s+= "</p>"	
		# puts "------"
		# puts s	
		return s
	end

	DEFAULT_CSS = "<style type='text/css'>
		.exceler-title h3 {
		color : blue;
	}
	.exceler-table {
		border-collapse: collapse;
		background-color: #ccf;
		width : 100%;
		border : 1px solid #888;
	}
	.exceler-table tr {
		border : 1px solid #888;
	}
	.exceler-table td {
		border : 1px solid #888;
	}
	</style>"

	#
	# 作成したHTMLファイルコンテンツにCSSを埋め込んで保存します。
	# ==== Args
	# file :: output file
	# content :: Content of the file
	# ==== Return
	# HTML Content
	#
	def self.write_html_file( file , content , css )
		f = open( file , "w" )
		if( css == nil )
			css = DEFAULT_CSS
		end
		f.puts( DEFAULT_CSS )
		f.puts( content )
		f.flush()
		f.close()
	end
	
	#
	# マップオブジェクトをCSVをにして保存します。
	# ==== Args
	# file :: output file
	# content :: Content of the file
	# ==== Return
	# HTML Content
	#
	def self.write_csv_file( file , map )
		keys = map.keys
		ks = ""
		vs = ""
		for key in keys
				ks += (key	 + ",")
				vs += (map[key].to_s + "," )
		end
		f = open( file , "w" )
		f.puts( ks )
		f.puts( vs )
		f.flush()
		f.close()
	end
	
end
