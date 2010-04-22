require 'bigdecimal'
require 'builder'
require 'cgi'

class Numeric
  # To simplify usage outside of rails
  unless instance_methods.include?('to_d')
    def to_d
      BigDecimal.new(self.to_s)
    end
  end
end

class ExcelXML

  class ExcelXMLBuilder
    def build
      @xml = Builder::XmlMarkup.new(:indent=>2)

      # The default implementation turns non-ASCII characters into entities, which does not work for Excel XML.
      def @xml._escape(text)
        CGI.escapeHTML(text)
      end
       
      @xml.instruct! :xml, :version=>"1.0", :encoding=>"UTF-8"
      @xml.Workbook({
        'xmlns'      => "urn:schemas-microsoft-com:office:spreadsheet",
        'xmlns:o'    => "urn:schemas-microsoft-com:office:office",
        'xmlns:x'    => "urn:schemas-microsoft-com:office:excel",
        'xmlns:html' => "http://www.w3.org/TR/REC-html40",
        'xmlns:ss'   => "urn:schemas-microsoft-com:office:spreadsheet"
      }) do
        @xml.Styles do
          @xml.Style "ss:ID" => "string" do
            @xml.NumberFormat "ss:Format" => "@"
          end
          @xml.Style "ss:ID" => "currency" do
            @xml.NumberFormat "ss:Format" => "#,##0.00"
          end          
        end
        yield self
      end
    end

    def <<(row)
      @xml.Row do
        row.each do |column|
          add_column(@xml, column)
        end
      end
    end
    
    def worksheet(name, opts = {})
      @xml.Worksheet 'ss:Name' => name.to_s do
        @xml.Table do
          opts[:column_widths] && opts[:column_widths].each { |width| @xml.Column('ss:Width' => width) }
          yield
        end
      end
    end

    private

    def add_column(xml, column)
      case column
      when Date
        xml.Cell { xml.Data column, 'ss:Type' => 'Date' }
      when BigDecimal
        xml.Cell('ss:StyleID' => 'currency') { xml.Data(column, 'ss:Type' => 'Number') }        
      when Numeric
        xml.Cell { xml.Data(column, 'ss:Type' => 'Number') }
      else
        xml.Cell('ss:StyleID' => 'string') { xml.Data(column.to_s, 'ss:Type' => 'String') }
      end
    end
  end

  def initialize(name="Sheet", data=[])
    @name = name
    @data = data
  end

  def self.build(&block)
    ExcelXMLBuilder.new.build(&block)
  end

  def to_sheet
    self.class.build do |e|
      e.worksheet @name do
        @data.each { |row| e << row }
      end
    end
  end

end
