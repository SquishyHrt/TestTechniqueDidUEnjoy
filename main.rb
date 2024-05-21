require_relative 'database_handling'
require_relative 'excel_reader'

excel_file_path = "Orders.xlsx"
reader = XLReader.new(excel_file_path)
db = DatabaseHandler.new("postgres") # Yes it is supposed to be due but time you know...

puts "Adding orders, packages and items to the database..."
db.add_all(reader.orders)

puts "Printing the orders from the the reader..."
puts reader.to_s

puts "Printing the items from the database..."
db.print_items

db.close