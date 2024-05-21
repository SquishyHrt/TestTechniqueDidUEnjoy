require_relative 'excel_reader'

require 'pg'

class DatabaseHandler
  def initialize(dbname)
    # Just connect to the database
    @conn = PG.connect(dbname: dbname)
  end

  # Add all the orders from the orders of XLReader to the database
  def add_order(orders)
    for i in 0..orders.length - 1
      tmp_name = "Order #{i + 1}" # default name of the excel file
      tmp_id = i + 1
      @conn.exec("INSERT INTO public.orders (orderid, odername) VALUES (#{tmp_id}, '#{tmp_name}')")
    end
  end

  # Add all the packages from the orders of XLReader to the database
  def add_packages(orders)
    for order in orders
      for package in order.packages
        tmp_package_id = order.order_id.to_s + package.id.to_s
        @conn.exec("INSERT INTO public.packages (packageid, orderid) VALUES (#{tmp_package_id.to_i}, #{order.order_id})")
      end
    end
  end

  # Add all the items from the orders of XLReader to the database
  def add_items(orders)
    for order in orders
      for package in order.packages
        package.items.each { |item|
          # Set the warranty here
          if item.warranty == nil || item.warranty == "NO" || item.warranty == ""
            item.warranty = false
          else
            item.warranty = true
          end
          @conn.exec("INSERT INTO public.items (itemid, name, price, ref, packageid, warranty, duration) VALUES (#{item.id}, '#{item.name}', #{item.price}, '#{item.ref}', #{item.package_id}, #{item.warranty}, #{item.duration})")
        }
      end
    end
  end

  # given a list of orders created by XLReader, add all the orders, packages and items to the database
  def add_all(orders)
    add_order(orders)
    add_packages(orders)
    add_items(orders)
  end

  def print_orders
    res = @conn.exec("SELECT * FROM public.orders")
    res.each do |row|
      puts row
    end
  end

  def print_packages
    res = @conn.exec("SELECT * FROM public.packages")
    res.each do |row|
      puts row
    end
  end

  def print_items
    res = @conn.exec("SELECT * FROM public.items")
    res.each do |row|
      puts row
    end
  end

  def close
    @conn.close
  end
end