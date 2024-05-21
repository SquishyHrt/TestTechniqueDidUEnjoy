require_relative 'database_handling'
require_relative 'excel_reader'

file_path = "Order.xlsx"
xl_reader = XLReader.new(file_path)