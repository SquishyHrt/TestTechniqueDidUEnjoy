require 'rubyXL'

workbook = RubyXL::Parser.parse("Orders.xlsx")

sheet =  workbook[0] # get first worksheet
puts sheet[0][0].value