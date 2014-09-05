# encoding: utf-8

module ExcelBook
  class Sheet
    attr_reader :sheet

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

    def delete
      sheet.Delete
    end
  end
end
