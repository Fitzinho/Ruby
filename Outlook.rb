require 'win32ole'
require 'csv'
require 'date'
require 'trollop'
#$outlook = WIN32OLE.new('Outlook.Application')
$VERBOSE=false
def getFolderByName(path)
	mapi = $outlook.GetNameSpace('MAPI')
	folder=mapi
	segments=[]
	path.sub(/^\\\\/,'').split("\\").each {|segment|
		segments <<segment
		segpath = "\\\\" << (segments.join("\\"))
		puts "Trying to access folder #{segpath}" if $VERBOSE
		folder=folder.Folders.Item(segment)
	}
	return folder
end
 
def main()
	
	opts = Trollop::options do 
		version "Outlook folder item header exporter"
  banner <<-EOS
Example usage:
       #{File.basename($0)} -p \"\\\\Public Folders\\Favorites\\Pentest\" -f 1/1/2012 -t 31/6/2012 -v -c mystats.csv       
EOS
		opt :path, "Full folder path", :type => :string, :required => true
		opt :from_date, "Earliest message date",:type => :string, :default=>"1/1/1970"
		opt :to_date, 	"Latest messgae date", :type=> :string, :default => Date.today().strftime("%d/%m/%Y")
		opt :csv,		"CSV file for output", :required => true, :type => :string
		opt :verbose,	"Verbose"
	end
	$outlook = WIN32OLE.new('Outlook.Application')
	$VERBOSE=opts[:verbose]
	properties="SentOn, SenderName, Subject,ConversationIndex, ConversationTopic".split(/\W+/)
	headers = properties.dup.concat ["WordCount","TotalWords"]
	fromtime=Time.new(*(opts[:from_date].split("/").reverse))
	totime=Time.new(*(opts[:to_date].split("/").reverse))
 
	CSV.open(opts[:csv],"wb",) do |csvfile|
		csvfile << headers
		folder=getFolderByName(opts[:path])
		row=nil
		i=0
		folder.Items.each {|m|
			i+=1
			
			if m.SentOn > totime then 
				puts "Skipping message No. #{i} as it was Sent On #{m.SentOn}" if $VERBOSE
				next
			end
			if m.SentOn < fromtime then
				break
			end
			row=[]
			properties.each{|p|
				begin
					row<<m.send(p)
				rescue NoMethodError
					row <<""
				end
			}
			lastname,firstname=m.SenderName.split(",")			
			body=m.Body
			endpos=body.index(lastname)
			if endpos.nil? then
				endpos=body.index(firstname)
			end
			if endpos.nil? then
				endpos=body.size
			end
			puts "Slice: #{endpos}/#{body.size}" if $VERBOSE
			beforeCrap=body.slice(0,endpos)
			worcount=beforeCrap.split.size
			row << worcount
			row << body.split.size			
			puts row.join(",") if  $VERBOSE			
			csvfile << row
		}
		
	end
end
 
if __FILE__ == $0
	main()
end
		segpath = "\\\\" << (segments.join("\\"))
		puts "Trying to access folder #{segpath}" if $VERBOSE
		folder=folder.Folders.Item(segment)
	}
	return folder
end
 
def main()
	
	opts = Trollop::options do 
		version "Outlook folder item header exporter"
  banner <<-EOS
Example usage:
       #{File.basename($0)} -p \"\\\\Public Folders\\Favorites\\Pentest\" -f 1/1/2012 -t 31/6/2012 -v -c mystats.csv       
EOS
		opt :path, "Full folder path", :type => :string, :required => true
		opt :from_date, "Earliest message date",:type => :string, :default=>"1/1/1970"
		opt :to_date, 	"Latest messgae date", :type=> :string, :default => Date.today().strftime("%d/%m/%Y")
		opt :csv,		"CSV file for output", :required => true, :type => :string
		opt :verbose,	"Verbose"
	end
	$outlook = WIN32OLE.new('Outlook.Application')
	$VERBOSE=opts[:verbose]
	properties="SentOn, SenderName, Subject,ConversationIndex, ConversationTopic".split(/\W+/)
	headers = properties.dup.concat ["WordCount","TotalWords"]
	fromtime=Time.new(*(opts[:from_date].split("/").reverse))
	totime=Time.new(*(opts[:to_date].split("/").reverse))
 
	CSV.open(opts[:csv],"wb",) do |csvfile|
		csvfile << headers
		folder=getFolderByName(opts[:path])
		row=nil
		i=0
		folder.Items.each {|m|
			i+=1
			
			if m.SentOn > totime then 
				puts "Skipping message No. #{i} as it was Sent On #{m.SentOn}" if $VERBOSE
				next
			end
			if m.SentOn < fromtime then
				break
			end
			row=[]
			properties.each{|p|
				begin
					row<<m.send(p)
				rescue NoMethodError
					row <<""
				end
			}
			lastname,firstname=m.SenderName.split(",")			
			body=m.Body
			endpos=body.index(lastname)
			if endpos.nil? then
				endpos=body.index(firstname)
			end
			if endpos.nil? then
				endpos=body.size
			end
			puts "Slice: #{endpos}/#{body.size}" if $VERBOSE
			beforeCrap=body.slice(0,endpos)
			worcount=beforeCrap.split.size
			row << worcount
			row << body.split.size			
			puts row.join(",") if  $VERBOSE			
			csvfile << row
		}
		
	end
end
 
if __FILE__ == $0
	main()
end