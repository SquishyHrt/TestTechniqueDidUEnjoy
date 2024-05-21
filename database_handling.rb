require_relative 'excel_reader'

require 'pg'

# public.items (
#     itemid integer NOT NULL,
#     name text,
#     price integer,
#     ref text,
#     packageid integer NOT NULL,
#     warranty boolean,
#     duration integer
# );

# public.orders (
#     orderid integer NOT NULL,
#     odername text NOT NULL
# );

# public.packages (
#     packageid integer NOT NULL,
#     orderid integer NOT NULL
# );

# first create an order the the package then the items

class DatabaseHandler
  def initialize(dbname)
    @conn = PG.connect(dbname: dbname)
  end
  
  def add_order(orders) # Orders is a list of order containing items
    for i in 0..orders.length - 1
      tmp_name = "Order #{i + 1}"
      tmp_id = i + 1
      @conn.exec("INSERT INTO public.orders (orderid, odername) VALUES (#{tmp_id}, '#{tmp_name}')")
    end
  end
  
  
end

db_handler = DatabaseHandler.new('postgres') #yes it should be due but hum...
xl_reader = XLReader.new("Order.xlsx")

db_handler.add_order(xl_reader.orders)