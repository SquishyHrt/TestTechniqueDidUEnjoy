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
    attr_accessor :nbCol, :nbRow
    def initialize(sheet, id)
        @orderId = id
        @sheet = sheet
        @nbRow = getNbRow
        @nbCol = getNbCol
        @items = []

        for i in 1..(@nbRow - 1)
            tmpPackage = @sheet[i][0].value.to_i
            tmpItemID = @sheet[i][1].value.to_i
            tmpLabel = @sheet[i][2].value
            tmpValue = @sheet[i][3].value

            if @items[tmpItemID] != nil
                case tmpLabel
                when "price"
                        @items[tmpItemID].price = tmpValue.to_i
                when "ref"
                        @items[tmpItemID].ref = tmpValue
                when "name"
                        @items[tmpItemID].name = tmpValue
                when "warranty"
                        @items[tmpItemID].warranty = tmpValue
                when "duration"
                        @items[tmpItemID].duration = (tmpValue != nil) ? tmpValue.to_i : 0
                else
                    puts "Error: unknown label #{tmpLabel}"
                end # end case
            else
                @items[tmpItemID] = Item.new(tmpItemID, tmpPackage)
            end # end if 
        end # end for
    end  # end init

    def getNbRow #lignes
        res = 0
        while @sheet[res] != nil
            res += 1
        end

        return res
    end

    def getNbCol
        res = 0
        while @sheet[0][res] != nil
            res += 1
        end

        res
    end

    def to_s
        res = "Order: #{@orderId}\n"
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
        @nbOrber = 0
        while @workbook[@nbOrber] != nil
            @orders.push(Order.new(@workbook[@nbOrber], @nbOrber))
            @nbOrber += 1
        end

    end

    def to_s
        res = ""
        @orders.each do |order|
            res += order.to_s + "\n"
        end
        return res
    end
end



reader = XLReader.new("Orders.xlsx")

puts reader