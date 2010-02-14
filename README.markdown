Library for creating Excel XML documents using Ruby.

About
=====

A simple Excel XML generating library extracted from the project I'm currently working on. Originally created by Henrik Nyh / Auktionskompaniet.com.

Dependencies
=====
    sudo gem install builder

Compability
=====
Has been tested with NeoOffice and MS Office 2003.

Usage
=====
A simple document:
    ExcelXML.new("People", [ [ "Name:", "Age:" ], [ "Jake", 42 ], [ "Kate", 33 ] ]).to_sheet

Getting excel_xml and generating a document:
    git clone git://github.com/joakimk/excel_xml.git
    cd excel_xml
    ruby -rubygems -rexcel_xml -e 'puts ExcelXML.new("People", [ [ "Name:", "Age:" ], [ "Jake", 42 ], [ "Kate", 33 ] ]).to_sheet' > doc.xml
    open doc.xml -a NeoOffice.app
    
Installing as a rails plugin:
    script/plugin install git://github.com/joakimk/excel_xml.git

Using multiple sheets in the same document:
    ExcelXML.build { |e|
      e.worksheet "People" do
        e << [ "Name:", "Age:" ]
        e << [ "Jake", 42 ]
        e << [ "Kate", 33 ]
      end
      
      e.worksheet "Books" do
        e << [ "Title:", "Year:" ]
        e << [ "The RSpec Book", 2010 ]
        e << [ "Extreme Programming Explained", 2004 ]
      end
    }
    
Using currency formatting:
    ExcelXML.new("Expenses", [ [ "Store:", "Amount:" ], [ "7-11", (9.50).to_d ], [ "IKEA", (350.25).to_d ] ]).to_sheet
    
Specifying column widths:
    ExcelXML.build { |e|
      e.worksheet "People", :column_widths => [ 200, 30 ] do
        e << [ "Name:", "Age:" ]
        e << [ "Jake", 42 ]
        e << [ "Kate", 33 ]
      end
    }

Running the specs
=====
    sudo gem install rspec nokogiri
    rake

Authors
====
###Contributors (alphabetical)
 - [Andreas Alin](http://github.com/aalin)
 - [Henrik Nyh](http://github.com/henrik)

Created by [Joakim Kolsjö](http://www.rubyblocks.se) for Auktionskompaniet.com.
joakim.kolsjo<$at$>gmail.com
