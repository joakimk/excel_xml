Library for creating Excel XML documents using Ruby.

About
=====

A simple Excel XML generating library originally created by Henrik Nyh. Extracted from
the project I'm currently working on.

Dependencies
=====
    sudo gem install nokogiri

Usage
=====
    -- Simple sheet --
    puts ExcelXML.new("People", [ [ "Name:", "Age:" ], [ "Jake", 42 ], [ "Kate", 33 ] ]).to_sheet
    
    -- Multiple sheets --
    puts ExcelXML.build { |e|
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
    
    -- Using currency formatting --
    puts ExcelXML.new("Expenses", [ [ "Store:", "Amount:" ], [ "7-11", (9.50).to_d ], [ "IKEA", (350.25).to_d ] ]).to_sheet
    
    -- Specifying column widths --
    puts ExcelXML.build { |e|
      e.worksheet "People", :column_widths => [ 200, 30 ] do
        e << [ "Name:", "Age:" ]
        e << [ "Jake", 42 ]
        e << [ "Kate", 33 ]
      end
    }

Running the specs
=====
    sudo gem install rspec
    rake

Authors
====
###Contributors (alphabetical)
 - [Andreas Alin](http://github.com/aalin)
 - [Henrik Nyh](http://github.com/henrik)

[Joakim Kolsjö](http://www.rubyblocks.se)
joakim.kolsjo<$at$>gmail.com

Hereby placed under public domain, do what you want, just do not hold me accountable...
