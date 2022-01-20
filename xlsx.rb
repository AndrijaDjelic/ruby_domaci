# gem install roo

require 'roo'

class Xlsx

    def initialize(rootFile)
      @excel = Roo::Spreadsheet.open(rootFile)
    end

  def findHeaders
  worksheets = @excel.sheets


  worksheets.each do |worksheet|
  puts ''
  puts worksheet
  puts '--------------'
  sheet = @excel.sheet(worksheet)

  
  if !sheet.nil?
    last_row    = sheet.last_row
    last_column = sheet.last_column

    first_row = nil
    first_col = nil
    flag = 0
   for row in 1..last_row do
     for col in 1..last_column do
      if !sheet.cell(row, col).nil?
        first_row = row
        first_col = col
        flag = 1
        break
      end
     end
     if flag==1
      break
     end
   end

    if !last_row.nil? and !last_column.nil?
      for col in first_col..last_column

        puts "["+ sheet.cell(first_row, col).to_s+"]"

        for row in (first_row+1)..last_row

          v = sheet.cell(row, col)
          
          if v.nil?
            puts "empty"
            else
            print sheet.cell(row, col).to_s + ', '
            end
          end
        end
        puts ''
      else
      puts 'Seems no data in sheet: ' + worksheet
      end
    end
    end
  end

  
end
