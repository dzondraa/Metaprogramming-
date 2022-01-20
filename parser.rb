require 'roo'


class Table 
    include Enumerable

    def elements
        @elements
    end

    def map
        @map 
    end
    
    def initialize(urlParam) 
        @map = {}
        @elements = createTable(urlParam)
    end
    
    def get(key)
        raise ArgumentError unless key.respond_to?(:to_s)    
        num = @map[key.to_s]
        num = num - 1
        return getColumn(num)
      end
    
    #   def set(key, value)
    #     raise ArgumentError unless key.respond_to?(:to_s)
    #     actual_key = key.to_s.downcase.to_sym
    #     @map[actual_key] = value
    #   end
    
      alias [] get
    #   alias []= set

    def row(nth)
        return @elements[nth]
    end

    def createTable(url)
        book = Roo::Spreadsheet.open url
        sheets = book.sheets
        table = Array.new 

        sheets.each do |worksheet| 
            count = 0
            columnNumber = 1;
            # Each row
            book.sheet(worksheet).each do |row|
                    temp = row.compact
                    if temp.length != 0 
                        if count == 0
                            row.each do |column|
                                map[column] = columnNumber  
                                columnNumber = columnNumber + 1

                                define_singleton_method(column) do
                                    self.get(column) 
                                  end                                
                                
                            end
                            count = count + 1
                        end
                        
                        table.push(temp)
                    end 
            end
        end
        return table
    end

    def each(&block)
        arr = Array.new
        @elements.each do |row|
          row.each do |element|
            block.call(element)
          end
        end
    end

    def getColumn(columnNumber) 
        tempArr = Array.new
        first = true
        @elements.each do |row|
            if(!first)
                tempArr.push(row[columnNumber])
            end
            first = false
        end
        return tempArr
    end

end


table = Table.new('C:\Users\v-dnikolic\Desktop\ruby\Book1.xls')
puts table.mesec2[0]

# puts table["mesec2"][0]