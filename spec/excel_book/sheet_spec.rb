require 'spec_helper'

describe ExcelBook do
  TEMP_DIR = File.expand_path('../../temp', Dir::pwd)

  it 'ExcelBook can create new sheet.' do
    ExcelBook::Book
  end
end
