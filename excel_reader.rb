require 'rubyXL'

class Item
  attr_accessor :id, :name, :price, :ref, :warranty, :duration, :package_id
  @@current_id = 0

  def initialize(package_id)
    @id = @@current_id
    @@current_id += 1
    @name = ""
    @price = 0
    @ref = ""
    @warranty = ""
    @duration = 0
    @package_id = package_id
  end

  def to_s
    "Item: #{@id}, name: #{@name}, price: #{@price}, ref: #{@ref}, warranty: #{@warranty}, duration: #{@duration}, package: #{@package_id}"
  end
end

class Package
  attr_accessor :items, :id

  def initialize(id)
    @items = []
    @id = id
  end
end

class Order
  attr_accessor :nb_col, :nb_row, :packages, :order_id

  def initialize(sheet, id)
    @order_id = id
    @sheet = sheet
    @nb_row = @sheet.count { |row| row != nil }
    @packages = [] # array of Package objects containing items

    for i in 1..(@nb_row - 1)
      tmp_package_id = @sheet[i][0].value.to_i
      tmp_item_id = @sheet[i][1].value.to_i
      tmp_label = @sheet[i][2].value
      tmp_value = @sheet[i][3].value

      if @packages[tmp_package_id] == nil
        tmp_package_id_concat = id.to_s + tmp_package_id.to_s # order_id + package_id
        @packages[tmp_package_id] = Package.new(tmp_package_id)
        @packages[tmp_package_id].items[tmp_item_id] = Item.new(tmp_package_id_concat)
      end

      if @packages[tmp_package_id].items[tmp_item_id] == nil # if no item in package so create it
        tmp_package_id_concat = id.to_s + tmp_package_id.to_s # order_id + package_id
        @packages[tmp_package_id].items[tmp_item_id] = Item.new(tmp_package_id_concat)
      end

      case tmp_label
      when "price"
        @packages[tmp_package_id].items[tmp_item_id].price = tmp_value.to_i
      when "ref"
        @packages[tmp_package_id].items[tmp_item_id].ref = tmp_value
      when "name"
        @packages[tmp_package_id].items[tmp_item_id].name = tmp_value
      when "warranty"
        @packages[tmp_package_id].items[tmp_item_id].warranty = tmp_value
      when "duration"
        @packages[tmp_package_id].items[tmp_item_id].duration = (tmp_value != nil) ? tmp_value.to_i : 0
      else
        puts "Error: unknown label #{tmp_label}"
      end

    end
  end

  def to_s
    res = "Order: #{@order_id}\n"
    @packages.each do |package|
      res += "Package: #{package.id}\n"
      package.items.each do |item|
        res += " " + item.to_s + "\n"
      end
    end
    res
  end
end

class XLReader
  attr_accessor :orders

  def initialize(path)
    @path = path
    @workbook = RubyXL::Parser.parse("Orders.xlsx")

    @orders = []
    @nb_order = 0
    while @workbook[@nb_order] != nil
      @orders.push(Order.new(@workbook[@nb_order], @nb_order + 1))
      @nb_order += 1
    end
  end

  def to_s
    res = ""
    @orders.each do |order|
      res += order.to_s + "\n"
    end
    res
  end
end