# encoding: utf-8

module ExcelBook
  class Sheet
    attr_reader :sheet

    CONVERT_LIST = {
        'A' => 0, 'B' => 1, 'C' => 2, 'D' => 3, 'E' => 4, 'F' => 5, 'G' => 6, 'H' => 7, 'I' => 8, 'J' => 9,
        'K' => 'A', 'L' => 'B', 'M' => 'C', 'N' => 'D', 'O' => 'E', 'P' => 'F', 'Q' => 'G', 'R' => 'H',
        'S' => 'I', 'T' => 'J', 'U' => 'K', 'V' => 'L', 'W' => 'M', 'X' => 'N', 'Y' => 'O', 'Z' => 'P'
    }

    def initialize(sheet)
      @sheet = sheet
    end

    def name=(name)
      sheet.name = name
    end

    def name
      sheet.name
    end

    def index
      sheet.index
    end

    def [](index)
      if /^(?<col>[A-Za-z]+)(?<row>[0-9]*)$/ =~ index
        result = sheet.columns(col2int(col))

        if row != ''
          result.rows(row.to_i).value
        else
          result
        end
      else
        raise '予期しないパラメータです'
      end
    end

    def []=(index)
    end

    def delete
      sheet.Delete
    end

    private
    def col2int(col)
      col.chars.map{|ch| CONVERT_LIST[ch]}.join.to_i(26)+1
    end
  end
end
