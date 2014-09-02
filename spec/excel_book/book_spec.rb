require 'spec_helper'

describe ExcelBook do
  it 'ExcelBook has a version.' do
    expect(ExcelBook::VERSION).not_to be nil
  end

  it 'ExcelBook can create new excel file.' do
    book = ExcelBook::Book.new(filename: 'example.xlsx')
  end
    
  end
end
