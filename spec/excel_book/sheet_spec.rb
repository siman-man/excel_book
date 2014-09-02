require 'spec_helper'

describe ExcelBook do
  it 'ExcelBook can create new sheet.' do
    filepath = File.expand_path('add_sheet.xlsx', TEMP_DIR)
    FileUtils.rm_f(filepath) if File.exist?(filepath)

    book = ExcelBook::Book.open(filepath, mode: 'w')

    begin
      p book.sheet_names
      book.add_sheet('Add')
      p book.sheet_names
    ensure
      book.save
      book.close
    end
  end
end
