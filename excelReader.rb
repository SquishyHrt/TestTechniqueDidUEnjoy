require 'rubyXL'

class Item
  attr_accessor :id, :name, :price, :ref, :warranty, :duration

  def initialize(id, package)
    @id = id
    @name = ""
    @price = 0
    @ref = ""
    @warranty = ""
    @duration = 0
    @package = package
  end

  def to_s
    "Item: #{@id}, name: #{@name}, price: #{@price}, ref: #{@ref}, warranty: #{@warranty}, duration: #{@duration}"
  end
end

class Order
  attr_accessor :nb_col, :nb_row

  def initialize(sheet, id)
    @order_id = id
    @sheet = sheet
    @nb_row = get_nb_row
    @nb_col = get_nb_col
    @items = []

    for i in 1..(@nb_row - 1)
      tmp_package = @sheet[i][0].value.to_i
      tmp_item_id = @sheet[i][1].value.to_i
      tmp_label = @sheet[i][2].value
      tmp_value = @sheet[i][3].value

      if @items[tmp_item_id] != nil
        case tmp_label
        when "price"
          @items[tmp_item_id].price = tmp_value.to_i
        when "ref"
          @items[tmp_item_id].ref = tmp_value
        when "name"
          @items[tmp_item_id].name = tmp_value
        when "warranty"
          @items[tmp_item_id].warranty = tmp_value
        when "duration"
          @items[tmp_item_id].duration = (tmp_value != nil) ? tmp_value.to_i : 0
        else
          puts "Error: unknown label #{tmp_label}"
        end # end case
      else
        @items[tmp_item_id] = Item.new(tmp_item_id, tmp_package)
      end # end if 
    end # end for
  end

  # end init

  def get_nb_row
    res = 0
    while @sheet[res] != nil
      res += 1
    end

    res
  end

  def get_nb_col
    res = 0
    while @sheet[0][res] != nil
      res += 1
    end

    res
  end

  def to_s
    res = "Order: #{@order_id}\n"
    @items.each do |item|
      res += item.to_s + "\n"
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
      @orders.push(Order.new(@workbook[@nb_order], @nb_order))
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

reader = XLReader.new("Orders.xlsx")
puts reader.orders